// src/index.js
// Entrypoint for the bot app: starts the HTTP server, configures the
// BotFramework adapter and Teams command adapter, and exposes routes for
// incoming Teams messages and proactive sends from external services.
const restify = require("restify");
const { commandApp } = require("./internal/initialize");
const { TeamsBot } = require("./teamsBot");
const { sendMessageToTeams } = require("./messages/sendMessageToTeams");
const { adapter } = require("./internal/adapter");
const config = require("./internal/config");
const appInsights = require("applicationinsights");
const { logger, logLifecycle, runWithLogContext, elapsedMs } = require("./internal/logger");
const { buildRequestLogContext } = require("./internal/requestContext");
const { createLifecycle } = require("./internal/lifecycle");
const { live, ready, setReady } = require("./internal/health");

const lifecycle = createLifecycle({
    getClient: () => appInsights.defaultClient,
});
lifecycle.register();

// Polyfill `globalThis.crypto.randomUUID` for environments where web crypto
// isn't available (some Azure/SDK libs expect `globalThis.crypto.randomUUID`).
if (typeof globalThis.crypto === "undefined" || typeof globalThis.crypto.randomUUID !== "function") {
    try {
        const nodeCrypto = require("crypto");
        if (nodeCrypto && typeof nodeCrypto.randomUUID === "function") {
            globalThis.crypto = globalThis.crypto || {};
            globalThis.crypto.randomUUID = nodeCrypto.randomUUID.bind(nodeCrypto);
            logger.info("Polyfilled globalThis.crypto.randomUUID from node crypto", {
                event: "BOT_STARTED",
            });
        } else if (nodeCrypto && nodeCrypto.webcrypto) {
            globalThis.crypto = nodeCrypto.webcrypto;
            logger.info("Polyfilled globalThis.crypto with node webcrypto", {
                event: "BOT_STARTED",
            });
        }
    } catch (polyErr) {
        logger.warn("Could not polyfill globalThis.crypto.randomUUID", {
            event: "BOT_STARTUP_FAILED",
            error: polyErr,
        });
    }
}

// Credentials come from internal/config.js (MicrosoftApp* on Azure, BOT_* locally).
// Proactive sends and inbound turns share the CloudAdapter in internal/adapter.js so
// UserAssignedMSI bots (no password) and secret-based bots stay on one auth path.
const MICROSOFT_APP_ID = config.MicrosoftAppId;

// A UserAssignedMSI bot has no password - the managed identity supplies its own tokens.
if (!MICROSOFT_APP_ID) {
    lifecycle.onStartupFailed("Set MicrosoftAppId (or BOT_ID).");
    process.exit(1);
}
if (config.MicrosoftAppType !== "UserAssignedMSI" && !config.MicrosoftAppPassword) {
    lifecycle.onStartupFailed("Set MicrosoftAppPassword (or BOT_PASSWORD/SECRET_BOT_PASSWORD).");
    process.exit(1);
}

const aiConn =
    process.env.APPLICATIONINSIGHTS_CONNECTION_STRING ||
    process.env.APPINSIGHTS_CONNECTION_STRING ||
    process.env.APPINSIGHTS_INSTRUMENTATIONKEY;

if (aiConn) {
    appInsights
        .setup(aiConn)
        .setAutoCollectRequests(true)
        .setAutoCollectPerformance(true)
        .setAutoCollectExceptions(true)
        .setAutoCollectDependencies(true)
        .setAutoDependencyCorrelation(true)
        .setAutoCollectConsole(false)
        .setInternalLogging(true, true)
        .start();
} else {
    // When no connection string/key is provided, setup() will rely on env vars if present.
    appInsights
        .setup()
        .setAutoCollectRequests(true)
        .setAutoCollectConsole(false)
        .setInternalLogging(true, true)
        .start();
}

// Store every trace. The SDK samples only when this is below 100.
if (appInsights.defaultClient && appInsights.defaultClient.config) {
    appInsights.defaultClient.config.samplingPercentage = 100;
}

logLifecycle("Application Insights initialized", {
    event: "APPINSIGHTS_READY",
    hasConnectionString: Boolean(aiConn),
});

const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.get("/health/live", live);
server.get("/health/ready", ready);
server.listen(process.env.port || process.env.PORT || 3978, () => {
    setReady(true);
    lifecycle.onStarted({
        port: Number(process.env.port || process.env.PORT || 3978),
        botAppIdSuffix: MICROSOFT_APP_ID.slice(-8),
        botAppType: config.MicrosoftAppType,
    });
});

// Capture unhandled errors on turns so we log full stacks and send telemetry.
adapter.onTurnError = async (context, error) => {
    try {
        const status = error && (error.statusCode || error.status || (error.response && error.response.status));
        const authFailed = status === 401 || status === 403;
        logger.error(authFailed ? "Teams authentication failed" : "Unhandled error on turn", {
            event: authFailed ? "TEAMS_AUTH_FAILED" : "TEAMS_ACTIVITY_FAILED",
            status,
            durationMs: elapsedMs(),
            error,
        });
        if (appInsights && appInsights.defaultClient && typeof appInsights.defaultClient.trackException === "function") {
            appInsights.defaultClient.trackException({ exception: error });
        }
    } catch (logErr) {
        logger.error("Failed to report turn error", {
            event: "TEAMS_ACTIVITY_FAILED",
            error: logErr,
        });
    }
};

const teamsBot = new TeamsBot();

// Bot messages
server.post("/api/messages", async (req, res) => {
    await runWithLogContext(buildRequestLogContext(req.body, req.headers), async () => {
        const activity = req.body || {};
        logger.info("Teams activity received", {
            event: "TEAMS_ACTIVITY_RECEIVED",
            activityId: activity.id,
            activityType: activity.type,
            activityName: activity.name,
            text: activity.text,
            value: activity.value,
            from: activity.from,
            recipient: activity.recipient,
            conversation: activity.conversation,
            channelData: activity.channelData,
            timestamp: activity.timestamp,
        });
        // Let TeamsFx ConversationBot process the request; optionally run TeamsBot logic inside the callback.
        await commandApp.requestHandler(req, res, async (turnContext) => {
            await teamsBot.run(turnContext);
        });
    });
});

// Proactive send (server generates its own BF token; client bearer ignored)
server.post("/api/sendToTeams", async (req, res) => {
    await runWithLogContext(buildRequestLogContext(req.body, req.headers), async () => {
    try {
        logger.info("Request received", {
            event: "TEAMS_PROACTIVE_RECEIVED",
            body: req.body,
        });
        const { conversationReference, cardContent, messageText, activityId, isUser, updateMode } = req.body || {};
        if (!conversationReference || (!cardContent && !messageText)) {
            logger.warn("Proactive send rejected", {
                event: "TEAMS_RESPONSE_FAILED",
                durationMs: elapsedMs(),
                body: req.body,
            });
            return res.send(400, { error: "Provide conversationReference and messageText or cardContent" });
        }
        const details = await sendMessageToTeams(
            adapter,
            conversationReference,
            cardContent,
            messageText,
            activityId,
            isUser,
            updateMode
        );
        logger.info("Proactive send completed", {
            event: "TEAMS_RESPONSE_SUCCEEDED",
            durationMs: elapsedMs(),
            details,
        });
        res.send(200, { status: "Message processed successfully", details });
    } catch (err) {
        const status = err && (err.statusCode || err.status || (err.response && err.response.status));
        const authFailed = status === 401 || status === 403;
        logger.error(authFailed ? "Teams authentication failed" : "Error sending message to Teams", {
            event: authFailed ? "TEAMS_AUTH_FAILED" : "TEAMS_RESPONSE_FAILED",
            status,
            durationMs: elapsedMs(),
            error: err,
        });
        try {
            if (appInsights && appInsights.defaultClient && typeof appInsights.defaultClient.trackException === "function") {
                appInsights.defaultClient.trackException({ exception: err });
            }
        } catch (aiErr) {
            logger.error("Failed to report error to App Insights", {
                event: "TEAMS_RESPONSE_FAILED",
                error: aiErr,
            });
        }
        res.send(err.statusCode || 500, err.body || err.message || "Internal Server Error");
    }
    });
});
