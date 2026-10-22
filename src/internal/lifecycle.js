// src/internal/lifecycle.js
// Startup, shutdown, and crash signals for the Windows App Service process.
// Flush uses the existing Application Insights client and is capped so a hung
// flush cannot keep the process alive. OOM and hard kill never reach these hooks.
const { logger, getInstanceId, getPackageVersion } = require("./logger");

const FLUSH_TIMEOUT_MS = 2000;

function defaultGetClient() {
    try {
        return require("applicationinsights").defaultClient;
    } catch {
        return null;
    }
}

/**
 * Flush pending telemetry. Resolves when the client callback runs, when flush
 * throws, or when the timeout elapses. A missing client resolves immediately.
 * @param {object|null|undefined} client
 * @param {{ timeoutMs?: number }} [options]
 * @returns {Promise<void>}
 */
function flushTelemetry(client, options = {}) {
    const timeoutMs = options.timeoutMs === undefined ? FLUSH_TIMEOUT_MS : options.timeoutMs;
    return new Promise((resolve) => {
        let settled = false;
        const finish = () => {
            if (settled) {
                return;
            }
            settled = true;
            resolve();
        };
        const timer = setTimeout(finish, timeoutMs);
        try {
            if (!client || typeof client.flush !== "function") {
                clearTimeout(timer);
                finish();
                return;
            }
            client.flush({
                isAppCrashing: true,
                callback: () => {
                    clearTimeout(timer);
                    finish();
                },
            });
        } catch {
            clearTimeout(timer);
            finish();
        }
    });
}

/**
 * @param {{
 *   exit?: (code: number) => void,
 *   getClient?: () => object|null|undefined,
 *   timeoutMs?: number,
 *   logInfo?: Function,
 *   logError?: Function,
 * }} [deps]
 */
function createLifecycle(deps = {}) {
    const exit = deps.exit || process.exit;
    const getClient = deps.getClient || defaultGetClient;
    const timeoutMs = deps.timeoutMs;
    const logInfo = deps.logInfo || ((message, fields) => logger.info(message, fields));
    const logError = deps.logError || ((message, fields) => logger.error(message, fields));
    let stopping = false;

    function onStarted(extra = {}) {
        logInfo("App listening", {
            event: "BOT_STARTED",
            pid: process.pid,
            nodeVersion: process.version,
            version: getPackageVersion(),
            instanceId: getInstanceId(),
            hosting: "azure-app-service-windows",
            startTime: new Date().toISOString(),
            ...extra,
        });
    }

    function onStartupFailed(message, fields = {}) {
        logError(message, {
            event: "BOT_STARTUP_FAILED",
            ...fields,
        });
    }

    async function shutdown(signal) {
        if (stopping) {
            return;
        }
        stopping = true;
        try {
            logInfo("Process stopping", {
                event: "BOT_STOPPING",
                signal,
                pid: process.pid,
                instanceId: getInstanceId(),
            });
        } catch {
            // Still flush and exit.
        }
        await flushTelemetry(safeClient(), { timeoutMs });
        exit(0);
    }

    async function onUncaughtException(error) {
        if (stopping) {
            return;
        }
        stopping = true;
        try {
            logError("Unhandled exception", {
                event: "BOT_UNHANDLED_EXCEPTION",
                error,
                pid: process.pid,
                instanceId: getInstanceId(),
            });
        } catch {
            // Still flush and exit. A listener that continues would hide the crash.
        }
        try {
            const client = safeClient();
            if (client && typeof client.trackException === "function") {
                const exception = error instanceof Error ? error : new Error(String(error));
                client.trackException({ exception });
            }
        } catch {
            // Telemetry must not block the exit.
        }
        await flushTelemetry(safeClient(), { timeoutMs });
        exit(1);
    }

    function onUnhandledRejection(reason) {
        try {
            logError("Unhandled rejection", {
                event: "BOT_UNHANDLED_REJECTION",
                error: reason,
                pid: process.pid,
                instanceId: getInstanceId(),
            });
        } catch {
            // Logging a rejection must not exit or throw into the process.
        }
    }

    function safeClient() {
        try {
            return getClient();
        } catch {
            return null;
        }
    }

    function register(target = process) {
        target.on("SIGTERM", () => {
            shutdown("SIGTERM");
        });
        target.on("SIGINT", () => {
            shutdown("SIGINT");
        });
        target.on("uncaughtException", (error) => {
            onUncaughtException(error);
        });
        target.on("unhandledRejection", (reason) => {
            onUnhandledRejection(reason);
        });
    }

    return {
        onStarted,
        onStartupFailed,
        shutdown,
        onUncaughtException,
        onUnhandledRejection,
        register,
    };
}

module.exports = {
    FLUSH_TIMEOUT_MS,
    flushTelemetry,
    createLifecycle,
};
