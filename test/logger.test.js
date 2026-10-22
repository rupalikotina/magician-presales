// test/logger.test.js
const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const {
    createLogger,
    runWithLogContext,
    getLogContext,
    toSchemaRecord,
    sanitizeError,
} = require("../src/internal/logger");
const {
    redactFields,
    sanitizeError: redactSanitizeError,
    prepareIdentifiers,
    sanitizeUrl,
} = require("../src/internal/redact");

function createCapture() {
    const stdoutChunks = [];
    const stderrChunks = [];
    return {
        stdoutChunks,
        stderrChunks,
        writeStdout: (line) => {
            stdoutChunks.push(String(line));
        },
        writeStderr: (line) => {
            stderrChunks.push(String(line));
        },
        allText() {
            return stdoutChunks.join("") + stderrChunks.join("");
        },
        lastJson() {
            const text = this.allText().trim();
            const lines = text.split("\n").filter((line) => line.startsWith("{"));
            assert.ok(lines.length > 0, `expected JSON log lines, got: ${JSON.stringify(text)}`);
            return JSON.parse(lines[lines.length - 1]);
        },
    };
}

function waitForLog(loggerInstance) {
    return new Promise((resolve) => {
        const transport = loggerInstance.transports[0];
        transport.once("logged", () => resolve());
    });
}

describe("redact", () => {
    it("keeps message text, payloads, and card data, and redacts credentials", () => {
        const rawEvent = { activity: { text: "hi" } };
        rawEvent.self = rawEvent;
        const redacted = redactFields({
            messageText: "hello from Teams",
            text: "hello from Teams",
            from: { id: "user-1", name: "Ada" },
            to: { id: "bot-1", name: "Loops" },
            authorization: "Bearer abc",
            token: "tok-1",
            password: "pw",
            cardJson: '{"type":"AdaptiveCard"}',
            rawEvent,
            actionData: { verb: "like", comment: "useful" },
            payload: { messageText: "hello from Teams", accessToken: "should-hide" },
            body: { conversationReference: { conversation: { id: "19:thread" } } },
            responseData: { status: "ok", message: "card updated" },
            activityId: "act-1",
            safe: "ok",
            headers: { Authorization: "Bearer x" },
        });

        assert.equal(redacted.messageText, "hello from Teams");
        assert.equal(redacted.text, "hello from Teams");
        assert.deepEqual(redacted.from, { id: "user-1", name: "Ada" });
        assert.deepEqual(redacted.to, { id: "bot-1", name: "Loops" });
        assert.equal(redacted.cardJson, '{"type":"AdaptiveCard"}');
        assert.equal(redacted.rawEvent.activity.text, "hi");
        assert.equal(redacted.rawEvent.self, "[circular]");
        assert.deepEqual(redacted.actionData, { verb: "like", comment: "useful" });
        assert.equal(redacted.payload.messageText, "hello from Teams");
        assert.equal(redacted.payload.accessToken, "should-hide");
        assert.equal(redacted.body.conversationReference.conversation.id, "19:thread");
        assert.deepEqual(redacted.responseData, { status: "ok", message: "card updated" });
        assert.equal(redacted.authorization, "Bearer abc");
        assert.equal(redacted.token, "tok-1");
        assert.equal(redacted.password, "pw");
        assert.deepEqual(redacted.headers, { Authorization: "Bearer x" });
        assert.equal(redacted.activityId, "act-1");
        assert.equal(redacted.safe, "ok");
    });

    it("sanitizes errors without headers or bodies", () => {
        const err = new Error("request failed");
        err.code = "ERR_BAD_RESPONSE";
        err.status = 500;
        err.config = {
            url: "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events?token=secret",
            headers: { Authorization: "Bearer leak" },
        };
        err.response = {
            status: 500,
            data: { detail: "backend said no", token: "hide-me" },
        };
        err.stack = "Error: request failed\n    at foo\nAuthorization: Bearer leak";

        const sanitized = redactSanitizeError(err);
        assert.equal(sanitized.name, "Error");
        assert.equal(sanitized.message, "request failed");
        assert.equal(sanitized.code, "ERR_BAD_RESPONSE");
        assert.equal(sanitized.status, 500);
        assert.equal(
            sanitized.url,
            "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events?token=secret"
        );
        assert.equal(sanitized.headers, undefined);
        assert.equal(sanitized.response, undefined);
        assert.equal(sanitized.config, undefined);
        assert.deepEqual(sanitized.responseData, { detail: "backend said no", token: "hide-me" });
        assert.ok(String(sanitized.stack).includes("Bearer"));
    });

    it("strips query strings from URLs", () => {
        assert.equal(
            sanitizeUrl("https://example.com/path?a=1&b=2"),
            "https://example.com/path"
        );
    });

    it("stores raw tenant and conversation identifiers by default", () => {
        const previous = process.env.LOG_RAW_IDENTIFIERS;
        delete process.env.LOG_RAW_IDENTIFIERS;
        try {
            const prepared = prepareIdentifiers("tenant-guid", "19:conv@thread.tacv2");
            assert.equal(prepared.tenantId, "tenant-guid");
            assert.equal(prepared.conversationId, "19:conv@thread.tacv2");
        } finally {
            if (previous === undefined) {
                delete process.env.LOG_RAW_IDENTIFIERS;
            } else {
                process.env.LOG_RAW_IDENTIFIERS = previous;
            }
        }
    });

    it("keeps raw identifiers when LOG_RAW_IDENTIFIERS=true", () => {
        const previous = process.env.LOG_RAW_IDENTIFIERS;
        process.env.LOG_RAW_IDENTIFIERS = "true";
        try {
            const prepared = prepareIdentifiers("tenant-guid", "19:conv@thread.tacv2");
            assert.equal(prepared.tenantId, "tenant-guid");
            assert.equal(prepared.conversationId, "19:conv@thread.tacv2");
        } finally {
            if (previous === undefined) {
                delete process.env.LOG_RAW_IDENTIFIERS;
            } else {
                process.env.LOG_RAW_IDENTIFIERS = previous;
            }
        }
    });
});

describe("toSchemaRecord", () => {
    it("emits the expected JSON shape and omits empty optional fields", () => {
        const previousRaw = process.env.LOG_RAW_IDENTIFIERS;
        delete process.env.LOG_RAW_IDENTIFIERS;
        try {
            const record = toSchemaRecord({
                level: "info",
                message: "Teams activity received",
                timestamp: "2026-01-01T00:00:00.000Z",
                event: "TEAMS_ACTIVITY_RECEIVED",
                correlationId: "corr-1",
                activityId: "act-1",
                tenantId: "tenant-1",
                conversationId: "conv-1",
                durationMs: 12,
                emptyShouldOmit: "",
                nullShouldOmit: null,
            });

            assert.equal(record.timestamp, "2026-01-01T00:00:00.000Z");
            assert.equal(record.level, "info");
            assert.equal(record.service, "teams-bot");
            assert.ok(record.environment);
            assert.ok(record.instanceId);
            assert.equal(record.event, "TEAMS_ACTIVITY_RECEIVED");
            assert.equal(record.correlationId, "corr-1");
            assert.equal(record.activityId, "act-1");
            assert.equal(record.tenantId, "tenant-1");
            assert.equal(record.conversationId, "conv-1");
            assert.equal(record.message, "Teams activity received");
            assert.equal(record.durationMs, 12);
            assert.equal(Object.prototype.hasOwnProperty.call(record, "emptyShouldOmit"), false);
            assert.equal(Object.prototype.hasOwnProperty.call(record, "nullShouldOmit"), false);
        } finally {
            if (previousRaw === undefined) {
                delete process.env.LOG_RAW_IDENTIFIERS;
            } else {
                process.env.LOG_RAW_IDENTIFIERS = previousRaw;
            }
        }
    });
});

describe("logger levels and App Insights transport", () => {
    it("respects LOG_LEVEL filtering", async () => {
        const capture = createCapture();
        const loggerInstance = createLogger({
            level: "warn",
            getClient: () => ({ trackTrace: () => {} }),
            writeStdout: capture.writeStdout,
            writeStderr: capture.writeStderr,
        });

        const logged = waitForLog(loggerInstance);
        loggerInstance.info("should be filtered", { event: "TEST_INFO" });
        loggerInstance.warn("should appear", { event: "TEST_WARN" });
        await logged;

        assert.ok(!capture.allText().includes("should be filtered"));
        assert.ok(capture.allText().includes("should appear"));
        const parsed = capture.lastJson();
        assert.equal(parsed.level, "warn");
        assert.equal(parsed.event, "TEST_WARN");
    });

    it("calls trackTrace on a stub client", async () => {
        const capture = createCapture();
        const traces = [];
        const loggerInstance = createLogger({
            level: "info",
            getClient: () => ({
                trackTrace: (payload) => traces.push(payload),
            }),
            writeStdout: capture.writeStdout,
            writeStderr: capture.writeStderr,
        });

        const logged = waitForLog(loggerInstance);
        loggerInstance.info("trace me", {
            event: "TEAMS_ACTIVITY_RECEIVED",
            correlationId: "c-1",
            activityId: "a-1",
        });
        await logged;
        await new Promise((resolve) => setImmediate(resolve));

        assert.equal(traces.length, 1);
        assert.equal(traces[0].message, "trace me");
        assert.equal(traces[0].severity, 1);
        assert.equal(traces[0].properties.event, "TEAMS_ACTIVITY_RECEIVED");
        assert.equal(traces[0].properties.correlationId, "c-1");
        assert.equal(traces[0].properties.activityId, "a-1");
        assert.equal(traces[0].properties.service, "teams-bot");
    });

    it("still writes JSON when trackTrace throws", async () => {
        const capture = createCapture();
        const loggerInstance = createLogger({
            level: "info",
            getClient: () => ({
                trackTrace: () => {
                    throw new Error("telemetry down");
                },
            }),
            writeStdout: capture.writeStdout,
            writeStderr: capture.writeStderr,
        });

        const logged = waitForLog(loggerInstance);
        assert.doesNotThrow(() => {
            loggerInstance.error("Loops webhook failed", {
                event: "LOOPS_WEBHOOK_FAILED",
                correlationId: "c-2",
            });
        });
        await logged;
        await new Promise((resolve) => setImmediate(resolve));

        const parsed = capture.lastJson();
        assert.equal(parsed.event, "LOOPS_WEBHOOK_FAILED");
        assert.equal(parsed.correlationId, "c-2");
        assert.equal(parsed.message, "Loops webhook failed");
    });

    it("still writes JSON when the App Insights client is missing", async () => {
        const capture = createCapture();
        const loggerInstance = createLogger({
            level: "info",
            getClient: () => null,
            writeStdout: capture.writeStdout,
            writeStderr: capture.writeStderr,
        });

        const logged = waitForLog(loggerInstance);
        loggerInstance.info("no client", { event: "BOT_STARTED" });
        await logged;

        const parsed = capture.lastJson();
        assert.equal(parsed.event, "BOT_STARTED");
        assert.equal(parsed.service, "teams-bot");
    });
});

describe("correlation context", () => {
    it("propagates correlationId through runWithLogContext", () => {
        assert.deepEqual(getLogContext(), {});

        runWithLogContext({ correlationId: "corr-outer", activityId: "act-1" }, () => {
            assert.equal(getLogContext().correlationId, "corr-outer");
            assert.equal(getLogContext().activityId, "act-1");

            runWithLogContext({ processingStatus: "webhook_accepted" }, () => {
                const ctx = getLogContext();
                assert.equal(ctx.correlationId, "corr-outer");
                assert.equal(ctx.activityId, "act-1");
                assert.equal(ctx.processingStatus, "webhook_accepted");
            });
        });

        assert.deepEqual(getLogContext(), {});
    });

    it("includes AsyncLocalStorage context in schema records", () => {
        runWithLogContext({ correlationId: "from-context", activityId: "act-ctx" }, () => {
            const record = toSchemaRecord({
                level: "info",
                message: "nested",
                timestamp: "2026-01-01T00:00:00.000Z",
                event: "TEAMS_ACTIVITY_RECEIVED",
            });
            assert.equal(record.correlationId, "from-context");
            assert.equal(record.activityId, "act-ctx");
        });
    });
});

describe("sanitizeError export", () => {
    it("matches redact sanitizeError for Error objects", () => {
        const err = new Error("boom");
        err.code = "ECONNABORTED";
        assert.deepEqual(sanitizeError(err).message, redactSanitizeError(err).message);
        assert.equal(sanitizeError(err).code, "ECONNABORTED");
    });
});
