// src/internal/redact.js
// Strip credentials from log fields. Message text, card payloads, activity
// participants, raw events, and backend responses are kept for debugging.
const crypto = require("crypto");

const CREDENTIAL_KEY_PATTERN =
    /^(authorization|password|secret|token|cookie|cookies|accessToken|refreshToken|clientSecret|MicrosoftAppPassword|BOT_PASSWORD|SECRET_BOT_PASSWORD|appPassword)$/i;

const MAX_DEPTH = 20;

/**
 * @param {unknown} value
 * @returns {string|undefined}
 */
function hashIdentifier(value) {
    if (value === undefined || value === null || value === "") {
        return undefined;
    }
    return crypto.createHash("sha256").update(String(value)).digest("hex").slice(0, 12);
}

function shouldUseRawIdentifiers() {
    // Store real tenant and conversation ids unless hashing is explicitly requested.
    return process.env.LOG_RAW_IDENTIFIERS !== "false";
}

function redactionEnabled() {
    // Credential redaction is off until LOG_REDACT=true. Full payloads are stored for now.
    return process.env.LOG_REDACT === "true";
}

/**
 * Strip query string from a URL. Returns the original string (minus query) if URL parsing fails.
 * @param {unknown} url
 * @returns {string|undefined}
 */
function sanitizeUrl(url) {
    if (url === undefined || url === null || url === "") {
        return undefined;
    }
    const raw = String(url);
    try {
        const parsed = new URL(raw);
        parsed.search = "";
        parsed.hash = "";
        return parsed.toString();
    } catch {
        return raw.split("?")[0];
    }
}

/**
 * Keep only safe diagnostic fields from an Error / axios-like error.
 * @param {unknown} error
 * @returns {object|undefined}
 */
function sanitizeError(error) {
    if (error === undefined || error === null) {
        return undefined;
    }
    if (typeof error === "string") {
        return { message: error };
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

    const status =
        error.status ||
        error.statusCode ||
        (error.response && error.response.status);
    if (status !== undefined && status !== null) {
        result.status = status;
    }

    const url =
        (error.config && error.config.url) ||
        error.url ||
        (error.request && error.request.path);
    const sanitized = redactionEnabled() ? sanitizeUrl(url) : (url ? String(url) : undefined);
    if (sanitized) {
        result.url = sanitized;
    }

    if (error.stack) {
        const lines = String(error.stack).split("\n");
        result.stack = (redactionEnabled()
            ? lines.filter((line) => !/authorization|bearer|password|secret|token/i.test(line))
            : lines
        ).slice(0, 20).join("\n");
    }

    if (error.response && error.response.data !== undefined) {
        result.responseData = redactFields(error.response.data);
    }

    return result;
}

/**
 * Recursively replace credential keys. Message bodies, payloads, from/to,
 * raw events, and response data are preserved. Does not mutate the input.
 * @param {unknown} value
 * @param {number} [depth]
 * @param {WeakSet<object>} [seen]
 * @returns {unknown}
 */
function redactFields(value, depth = 0, seen = new WeakSet()) {
    if (value === undefined || value === null) {
        return value;
    }
    if (typeof value !== "object") {
        return value;
    }
    if (depth > MAX_DEPTH) {
        return "[truncated]";
    }
    if (seen.has(value)) {
        return "[circular]";
    }
    seen.add(value);

    if (Array.isArray(value)) {
        return value.map((item) => redactFields(item, depth + 1, seen));
    }
    if (value instanceof Error) {
        return sanitizeError(value);
    }

    const out = {};
    for (const [key, child] of Object.entries(value)) {
        if (redactionEnabled() && CREDENTIAL_KEY_PATTERN.test(key)) {
            out[key] = "[redacted]";
            continue;
        }
        if (redactionEnabled() && /header/i.test(key) && child && typeof child === "object") {
            out[key] = "[redacted]";
            continue;
        }
        out[key] = redactFields(child, depth + 1, seen);
    }
    return out;
}

/**
 * Prepare tenant / conversation identifiers for logging.
 * @param {string|undefined} tenantId
 * @param {string|undefined} conversationId
 * @returns {{ tenantId?: string, conversationId?: string }}
 */
function prepareIdentifiers(tenantId, conversationId) {
    const raw = shouldUseRawIdentifiers();
    const result = {};
    if (tenantId !== undefined && tenantId !== null && tenantId !== "") {
        result.tenantId = raw ? String(tenantId) : hashIdentifier(tenantId);
    }
    if (conversationId !== undefined && conversationId !== null && conversationId !== "") {
        result.conversationId = raw ? String(conversationId) : hashIdentifier(conversationId);
    }
    return result;
}

module.exports = {
    hashIdentifier,
    shouldUseRawIdentifiers,
    redactionEnabled,
    sanitizeUrl,
    sanitizeError,
    redactFields,
    prepareIdentifiers,
};
