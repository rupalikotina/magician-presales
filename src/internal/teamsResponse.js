// src/internal/teamsResponse.js
// Log Teams connector send/update outcomes without changing the thrown error.
const { logger, elapsedMs } = require("./logger");

function statusOf(error) {
    if (!error) {
        return undefined;
    }
    return error.statusCode || error.status || (error.response && error.response.status);
}

function logFailure(error, activity) {
    const status = statusOf(error);
    const authFailed = status === 401 || status === 403;
    logger.error(authFailed ? "Teams authentication failed" : "Teams response failed", {
        event: authFailed ? "TEAMS_AUTH_FAILED" : "TEAMS_RESPONSE_FAILED",
        durationMs: elapsedMs(),
        status,
        activity,
        body: error && error.response && (error.response.data || error.response.body),
        error,
    });
}

async function sendTeamsActivity(context, activity) {
    try {
        const result = await context.sendActivity(activity);
        logger.info("Teams response sent", {
            event: "TEAMS_RESPONSE_SUCCEEDED",
            durationMs: elapsedMs(),
            activity,
            response: result,
        });
        return result;
    } catch (error) {
        logFailure(error, activity);
        throw error;
    }
}

async function updateTeamsActivity(context, activity) {
    try {
        const result = await context.updateActivity(activity);
        logger.info("Teams response updated", {
            event: "TEAMS_RESPONSE_SUCCEEDED",
            durationMs: elapsedMs(),
            activity,
            response: result,
        });
        return result;
    } catch (error) {
        logFailure(error, activity);
        throw error;
    }
}

function logConnectorError(message, err) {
    const status = statusOf(err);
    if (status === 401 || status === 403) {
        logger.error("Teams authentication failed", {
            event: "TEAMS_AUTH_FAILED",
            durationMs: elapsedMs(),
            status,
            error: err,
        });
        return;
    }
    logger.error(message, { event: "TEAMS_API_FAILED", status, error: err });
}

module.exports = {
    sendTeamsActivity,
    updateTeamsActivity,
    logConnectorError,
};
