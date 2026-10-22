// src/internal/requestContext.js
// One correlation id for an inbound HTTP request. Handlers must not mint another.
const crypto = require("crypto");

function headerValue(headers, name) {
    if (!headers) {
        return undefined;
    }
    const target = name.toLowerCase();
    for (const [key, value] of Object.entries(headers)) {
        if (key.toLowerCase() === target && value !== undefined && value !== null && value !== "") {
            return Array.isArray(value) ? value[0] : value;
        }
    }
    return undefined;
}

/**
 * @param {object} [body]
 * @param {object} [headers]
 * @returns {{ correlationId: string, activityId?: string, startedAt: number, traceparent?: string }}
 */
function buildRequestLogContext(body, headers) {
    const activity = body || {};
    const activityId = activity.id || activity.activityId || undefined;
    const context = {
        correlationId: activityId || crypto.randomUUID(),
        startedAt: Date.now(),
    };
    if (activityId) {
        context.activityId = activityId;
    }
    const traceparent = headerValue(headers, "traceparent");
    if (traceparent) {
        context.traceparent = traceparent;
    }
    return context;
}

module.exports = {
    headerValue,
    buildRequestLogContext,
};
