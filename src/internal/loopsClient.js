// src/internal/loopsClient.js
// Same axios POST as before, plus correlation headers. Does not record a
// second dependency; Application Insights HTTP auto-collection already does that.
const axios = require("axios");
const { getLogContext, logger } = require("./logger");
const { headerValue } = require("./requestContext");
const { sanitizeUrl } = require("./redact");

/**
 * Copy correlation onto outbound headers without replacing a traceparent
 * the caller already set. Omit traceparent when nobody has one so the
 * Application Insights HTTP instrumentation can add it.
 * @param {object} [headers]
 * @returns {object}
 */
function applyCorrelationHeaders(headers) {
    const next = { ...(headers || {}) };
    const ctx = getLogContext();

    if (!headerValue(next, "x-correlation-id") && ctx.correlationId) {
        next["x-correlation-id"] = ctx.correlationId;
    }
    if (!headerValue(next, "traceparent") && ctx.traceparent) {
        next.traceparent = ctx.traceparent;
    }
    return next;
}

function responseStatus(error) {
    if (!error) {
        return undefined;
    }
    if (error.statusCode || error.status) {
        return error.statusCode || error.status;
    }
    return error.response && error.response.status;
}

function dependencyTarget(url) {
    const clean = sanitizeUrl(url) || "";
    try {
        const parsed = new URL(clean);
        return `${parsed.host}${parsed.pathname}`;
    } catch {
        return String(clean).split("?")[0];
    }
}

function classifyLoopsError(error) {
    const code = error && (error.code || error.errno);
    if (code === "ECONNABORTED" || code === "ETIMEDOUT") {
        return "timeout";
    }
    if (responseStatus(error)) {
        return "http_error";
    }
    return "network_error";
}

function errorWithoutBody(error) {
    if (!error || typeof error !== "object") {
        return { message: error === undefined || error === null ? "" : String(error) };
    }
    const result = {};
    if (error.name) {
        result.name = String(error.name);
    }
    if (error.message) {
        result.message = String(error.message);
    }
    if (error.code) {
        result.code = String(error.code);
    }
    const status = responseStatus(error);
    if (status) {
        result.status = status;
    }
    return result;
}

function dependencyFields(url, status, durationMs, errorClass) {
    const fields = {
        dependency: "loops-webhook",
        method: "POST",
        target: dependencyTarget(url),
        durationMs,
        retryCount: 0,
        correlationId: getLogContext().correlationId,
    };
    if (status) {
        fields.status = status;
    }
    if (errorClass) {
        fields.class = errorClass;
    }
    return fields;
}

function responseBody(errorOrResponse) {
    if (!errorOrResponse) {
        return undefined;
    }
    if (errorOrResponse.data !== undefined && errorOrResponse.status) {
        return errorOrResponse.data;
    }
    const response = errorOrResponse.response;
    if (!response) {
        return undefined;
    }
    return response.data !== undefined ? response.data : response.body;
}

/**
 * POST to Loops with the same URL, body, and error as axios.post.
 * A successful HTTP response is webhook acceptance, not digital-worker completion.
 * @param {string} url
 * @param {unknown} data
 * @param {object} [config]
 * @param {{ post?: Function, logger?: { info: Function, error: Function } }} [deps]
 */
async function postLoops(url, data, config, deps) {
    const httpPost = (deps && deps.post) || axios.post;
    const log = (deps && deps.logger) || logger;
    const nextConfig = { ...(config || {}) };
    nextConfig.headers = applyCorrelationHeaders(config && config.headers);

    log.info("Loops webhook started", {
        event: "LOOPS_WEBHOOK_STARTED",
        dependency: "loops-webhook",
        method: "POST",
        target: dependencyTarget(url),
        retryCount: 0,
        correlationId: getLogContext().correlationId,
        payload: data,
    });

    const started = Date.now();
    try {
        const response = await httpPost(url, data, nextConfig);
        log.info("Loops webhook accepted", {
            event: "LOOPS_WEBHOOK_SUCCEEDED",
            processingStatus: "webhook_accepted",
            ...dependencyFields(url, response && response.status, Date.now() - started),
            payload: data,
            response: responseBody(response),
        });
        return response;
    } catch (error) {
        log.error("Loops webhook failed", {
            event: "LOOPS_WEBHOOK_FAILED",
            processingStatus: "webhook_failed",
            ...dependencyFields(url, responseStatus(error), Date.now() - started, classifyLoopsError(error)),
            payload: data,
            error: errorWithoutBody(error),
        });
        throw error;
    }
}

module.exports = {
    applyCorrelationHeaders,
    classifyLoopsError,
    dependencyTarget,
    postLoops,
};
