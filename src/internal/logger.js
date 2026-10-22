// src/internal/logger.js
// Structured JSON logger for the Teams bot. Writes to stdout/stderr (never
// console.*) and forwards traces through the existing Application Insights client.
const { AsyncLocalStorage } = require("async_hooks");
const os = require("os");
const path = require("path");
const winston = require("winston");
const Transport = require("winston-transport");
const {
    sanitizeError,
    redactFields,
    prepareIdentifiers,
} = require("./redact");

const SERVICE_NAME = "teams-bot";
const CONTEXT_KEYS = [
    "correlationId",
    "activityId",
    "tenantId",
    "conversationId",
    "event",
    "durationMs",
    "processingStatus",
];

const logContext = new AsyncLocalStorage();

const customLevels = {
    levels: {
        error: 0,
        warn: 1,
        info: 2,
        debug: 3,
    },
    colors: {
        error: "red",
        warn: "yellow",
        info: "green",
        debug: "blue",
    },
};

const SEVERITY = {
    debug: 0, // Verbose
    info: 1, // Information
    warn: 2, // Warning
    error: 3, // Error
};

function getEnvironment() {
    if (process.env.NODE_ENV) {
        return process.env.NODE_ENV;
    }
    if (process.env.RUNNING_ON_AZURE === "1") {
        return "production";
    }
    return "development";
}

function getInstanceId() {
    return process.env.WEBSITE_INSTANCE_ID || os.hostname();
}

function getPackageVersion() {
    try {
        // eslint-disable-next-line import/no-dynamic-require, global-require
        const pkg = require(path.join(__dirname, "..", "..", "package.json"));
        return pkg.version || "unknown";
    } catch {
        return "unknown";
    }
}

/**
 * Remove Winston Symbol keys and empty optional fields before emission.
 * @param {object} info
 * @returns {object}
 */
function toSchemaRecord(info) {
    const ctx = logContext.getStore() || {};
    const record = {
        timestamp: info.timestamp || new Date().toISOString(),
        level: info.level,
        service: SERVICE_NAME,
        environment: getEnvironment(),
        instanceId: getInstanceId(),
        message: info.message,
    };

    const merged = { ...ctx };
    for (const [key, value] of Object.entries(info)) {
        if (typeof key === "symbol") {
            continue;
        }
        if (
            key === "level" ||
            key === "message" ||
            key === "timestamp" ||
            key === "splat" ||
            key === "stack"
        ) {
            continue;
        }
        if (value !== undefined) {
            merged[key] = value;
        }
    }

    const identifiers = prepareIdentifiers(merged.tenantId, merged.conversationId);
    if (identifiers.tenantId) {
        record.tenantId = identifiers.tenantId;
    }
    if (identifiers.conversationId) {
        record.conversationId = identifiers.conversationId;
    }
    delete merged.tenantId;
    delete merged.conversationId;

    for (const key of CONTEXT_KEYS) {
        if (key === "tenantId" || key === "conversationId") {
            continue;
        }
        const value = merged[key];
        if (value !== undefined && value !== null && value !== "") {
            record[key] = value;
        }
        delete merged[key];
    }

    if (merged.error !== undefined) {
        record.error = sanitizeError(merged.error);
        delete merged.error;
    }

    const redacted = redactFields(merged);
    if (redacted && typeof redacted === "object") {
        for (const [key, value] of Object.entries(redacted)) {
            if (value !== undefined && value !== null && value !== "") {
                record[key] = value;
            }
        }
    }

    return record;
}

/**
 * Winston transport that writes one JSON line per log to stdout/stderr.
 * Does not call console.*, so Application Insights console capture will not
 * double-ingest these records. Tests may inject writeStdout/writeStderr.
 */
class StdioJsonTransport extends Transport {
    /**
     * @param {{ level?: string, writeStdout?: Function, writeStderr?: Function }} [opts]
     */
    constructor(opts = {}) {
        super(opts);
        this.writeStdout = opts.writeStdout || ((line) => process.stdout.write(line));
        this.writeStderr = opts.writeStderr || ((line) => process.stderr.write(line));
    }

    log(info, callback) {
        setImmediate(() => this.emit("logged", info));
        try {
            const record = toSchemaRecord(info);
            const line = `${JSON.stringify(record)}\n`;
            if (info.level === "error" || info.level === "warn") {
                this.writeStderr(line);
            } else {
                this.writeStdout(line);
            }
        } catch {
            // Never interrupt the caller for logging failures.
        }
        callback();
    }
}

/**
 * Forwards structured records to the existing Application Insights client via trackTrace.
 */
class AppInsightsTransport extends Transport {
    /**
     * @param {{ getClient?: () => object|null|undefined }} [opts]
     */
    constructor(opts = {}) {
        super(opts);
        this.getClient =
            opts.getClient ||
            (() => {
                try {
                    // Lazy require so tests can inject a stub without starting AI.
                    // eslint-disable-next-line global-require
                    const appInsights = require("applicationinsights");
                    return appInsights.defaultClient;
                } catch {
                    return null;
                }
            });
    }

    log(info, callback) {
        setImmediate(() => this.emit("logged", info));
        try {
            const client = this.getClient();
            if (client && typeof client.trackTrace === "function") {
                const record = toSchemaRecord(info);
                const properties = {};
                for (const [key, value] of Object.entries(record)) {
                    if (key === "message" || key === "timestamp") {
                        continue;
                    }
                    if (value === undefined || value === null) {
                        continue;
                    }
                    properties[key] =
                        typeof value === "object" ? JSON.stringify(value) : String(value);
                }
                client.trackTrace({
                    message: record.message || "",
                    severity: SEVERITY[info.level] ?? SEVERITY.info,
                    properties,
                });
            }
        } catch {
            // Telemetry must never interrupt Teams message processing.
        }
        callback();
    }
}

/**
 * @param {{
 *   level?: string,
 *   getClient?: () => object|null|undefined,
 *   silent?: boolean,
 *   writeStdout?: Function,
 *   writeStderr?: Function,
 * }} [options]
 * @returns {winston.Logger}
 */
function createLogger(options = {}) {
    // Default to debug so info, warn, error, and debug are all stored.
    // Operational detail is still logged at info so it shows up without a debug filter.
    const level = options.level || process.env.LOG_LEVEL || "debug";
    const transports = [];

    if (!options.silent) {
        transports.push(
            new StdioJsonTransport({
                level,
                writeStdout: options.writeStdout,
                writeStderr: options.writeStderr,
            })
        );
        transports.push(
            new AppInsightsTransport({
                level,
                getClient: options.getClient,
            })
        );
    }

    return winston.createLogger({
        levels: customLevels.levels,
        level,
        defaultMeta: {},
        format: winston.format.combine(
            winston.format.timestamp({ format: () => new Date().toISOString() }),
            winston.format.errors({ stack: true })
        ),
        transports,
        exitOnError: false,
    });
}

const logger = createLogger();

/**
 * Run fn with a request-scoped log context (correlationId, activityId, …).
 * @template T
 * @param {object} context
 * @param {() => T} fn
 * @returns {T}
 */
function runWithLogContext(context, fn) {
    const parent = logContext.getStore() || {};
    return logContext.run({ ...parent, ...context }, fn);
}

/**
 * @returns {object}
 */
function getLogContext() {
    return logContext.getStore() || {};
}

function getLogStore() {
    return logContext.getStore();
}

function elapsedMs() {
    const startedAt = getLogContext().startedAt;
    if (typeof startedAt !== "number") {
        return undefined;
    }
    return Date.now() - startedAt;
}

function logLifecycle(message, fields = {}) {
    logger.info(message, { event: fields.event || "BOT_STARTED", ...fields });
}

function logTeamsActivity(message, fields = {}) {
    logger.info(message, { event: fields.event || "TEAMS_ACTIVITY_RECEIVED", ...fields });
}

function logExternalCall(message, fields = {}) {
    const { level: fieldLevel, ...rest } = fields;
    const level = fieldLevel || (rest.error ? "error" : "info");
    logger.log(level, message, { event: rest.event || "EXTERNAL_CALL", ...rest });
}

function logAuthFailure(message, fields = {}) {
    logger.error(message, { event: fields.event || "TEAMS_AUTH_FAILED", ...fields });
}

function logLoopsForward(message, fields = {}) {
    const { level: fieldLevel, ...rest } = fields;
    const level = fieldLevel || (rest.error ? "error" : "info");
    logger.log(level, message, {
        event: rest.event || "LOOPS_WEBHOOK_STARTED",
        ...rest,
    });
}

function logDuplicate(message, fields = {}) {
    logger.warn(message, { event: fields.event || "TEAMS_DUPLICATE_ACTIVITY_DETECTED", ...fields });
}

function logUnhandled(message, fields = {}) {
    logger.error(message, { event: fields.event || "BOT_UNHANDLED_EXCEPTION", ...fields });
}

module.exports = {
    logger,
    createLogger,
    runWithLogContext,
    getLogContext,
    getLogStore,
    elapsedMs,
    logLifecycle,
    logTeamsActivity,
    logExternalCall,
    logAuthFailure,
    logLoopsForward,
    logDuplicate,
    logUnhandled,
    getEnvironment,
    getInstanceId,
    getPackageVersion,
    SERVICE_NAME,
    StdioJsonTransport,
    AppInsightsTransport,
    toSchemaRecord,
    sanitizeError,
};
