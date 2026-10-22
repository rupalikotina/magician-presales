// test/loopsClient.test.js
const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { runWithLogContext, toSchemaRecord, createLogger } = require("../src/internal/logger");
const { buildRequestLogContext } = require("../src/internal/requestContext");
const { applyCorrelationHeaders, postLoops } = require("../src/internal/loopsClient");

describe("request correlation", () => {
    it("uses the activity id as the correlation id on a nested log", () => {
        const context = buildRequestLogContext(
            { id: "activity-42", text: "hello from Teams" },
            { traceparent: "00-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa-bbbbbbbbbbbbbbbb-01" }
        );
        assert.equal(context.correlationId, "activity-42");
        assert.equal(context.activityId, "activity-42");
        assert.equal(typeof context.startedAt, "number");

        runWithLogContext(context, () => {
            const record = toSchemaRecord({
                level: "info",
                message: "nested",
                timestamp: "2026-01-01T00:00:00.000Z",
                event: "TEAMS_ACTIVITY_RECEIVED",
            });
            assert.equal(record.correlationId, "activity-42");
            assert.equal(record.activityId, "activity-42");
            assert.equal(record.event, "TEAMS_ACTIVITY_RECEIVED");
        });
    });

    it("mints a correlation id when the body has no activity id", () => {
        const context = buildRequestLogContext({ messageText: "proactive" }, {});
        assert.equal(typeof context.correlationId, "string");
        assert.ok(context.correlationId.length > 0);
        assert.equal(context.activityId, undefined);
        assert.equal(context.traceparent, undefined);
    });
});

describe("Loops correlation headers", () => {
    it("copies an inbound traceparent and does not replace one already set", () => {
        const inbound = "00-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa-bbbbbbbbbbbbbbbb-01";
        const existing = "00-cccccccccccccccccccccccccccccccc-dddddddddddddddd-01";
        const context = buildRequestLogContext({ id: "activity-42" }, { traceparent: inbound });

        runWithLogContext(context, () => {
            const copied = applyCorrelationHeaders({ "Content-Type": "application/json" });
            assert.equal(copied.traceparent, inbound);
            assert.equal(copied["x-correlation-id"], "activity-42");
            assert.equal(copied["Content-Type"], "application/json");

            const kept = applyCorrelationHeaders({
                traceparent: existing,
                "x-correlation-id": "already-set",
            });
            assert.equal(kept.traceparent, existing);
            assert.equal(kept["x-correlation-id"], "already-set");
        });
    });

    it("omits traceparent when the inbound request has none", () => {
        const context = buildRequestLogContext({ id: "activity-7" }, {});
        runWithLogContext(context, () => {
            const headers = applyCorrelationHeaders({});
            assert.equal(Object.prototype.hasOwnProperty.call(headers, "traceparent"), false);
            assert.equal(headers["x-correlation-id"], "activity-7");
        });
    });

    it("posts the same url and body and rejects with the same error", async () => {
        const body = JSON.stringify({ messageText: "hello", from: { id: "user-1" } });
        const calls = [];
        await runWithLogContext(buildRequestLogContext({ id: "activity-42" }, { traceparent: "00-aaa-bbb-01" }), async () => {
            const response = await postLoops(
                "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events",
                body,
                { headers: { "Content-Type": "application/json" } },
                {
                    logger: createLogger({ getClient: () => null, writeStdout: () => {}, writeStderr: () => {} }),
                    post: async (url, data, config) => {
                        calls.push({ url, data, config });
                        return { status: 200, data: { status: "ok", message: "stored" } };
                    },
                }
            );
            assert.equal(response.data.message, "stored");
        });

        assert.equal(calls.length, 1);
        assert.equal(calls[0].url, "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events");
        assert.equal(calls[0].data, body);
        assert.equal(calls[0].config.headers["Content-Type"], "application/json");
        assert.equal(calls[0].config.headers["x-correlation-id"], "activity-42");

        const backendError = new Error("backend down");
        await assert.rejects(
            () => postLoops("https://example.test/hook", body, undefined, {
                logger: createLogger({ getClient: () => null, writeStdout: () => {}, writeStderr: () => {} }),
                post: async () => {
                    throw backendError;
                },
            }),
            (caught) => caught === backendError
        );
    });

    it("does not record a second dependency", () => {
        const source = fs.readFileSync(path.join(__dirname, "..", "src", "internal", "loopsClient.js"), "utf8");
        assert.equal(source.includes("trackDependency"), false);
        assert.equal(source.includes("trackMetric"), false);
    });
});

describe("Loops webhook lifecycle", () => {
    function captureLogger() {
        const lines = [];
        const log = createLogger({
            level: "debug",
            getClient: () => null,
            writeStdout: (line) => lines.push(String(line)),
            writeStderr: (line) => lines.push(String(line)),
        });
        return {
            log,
            records() {
                return lines
                    .join("")
                    .trim()
                    .split("\n")
                    .filter((line) => line.startsWith("{"))
                    .map((line) => JSON.parse(line));
            },
        };
    }

    it("logs started and accepted with duration without changing the posted body", async () => {
        const body = JSON.stringify({ messageText: "hello from Teams", from: { id: "user-1" } });
        const posted = [];
        const capture = captureLogger();
        const responseBody = { status: "ok", cardJson: "{\"title\":\"card\"}" };

        await runWithLogContext(buildRequestLogContext({ id: "activity-42" }, {}), async () => {
            const response = await postLoops(
                "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events",
                body,
                { headers: { "Content-Type": "application/json" } },
                {
                    logger: capture.log,
                    post: async (url, data, config) => {
                        posted.push({ url, data, config });
                        return { status: 202, data: responseBody };
                    },
                }
            );
            assert.equal(response.status, 202);
            assert.equal(response.data, responseBody);
        });

        assert.equal(posted.length, 1);
        assert.equal(posted[0].url, "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events");
        assert.equal(posted[0].data, body);
        assert.equal(posted[0].config.headers["Content-Type"], "application/json");

        const records = capture.records();
        const started = records.find((record) => record.event === "LOOPS_WEBHOOK_STARTED");
        const accepted = records.find((record) => record.event === "LOOPS_WEBHOOK_SUCCEEDED");
        assert.equal(started.payload, body);
        assert.equal(accepted.processingStatus, "webhook_accepted");
        assert.equal(accepted.status, 202);
        assert.equal(typeof accepted.durationMs, "number");
        assert.ok(accepted.durationMs >= 0);
        assert.deepEqual(accepted.response, responseBody);
        assert.equal(accepted.payload, body);
    });

    it("logs failure with status and sanitized error and rethrows the same error", async () => {
        const body = JSON.stringify({ messageText: "hello", token: "request-secret" });
        const posted = [];
        const capture = captureLogger();
        const backendError = Object.assign(new Error("backend down"), {
            response: {
                status: 500,
                data: { token: "response-secret", detail: "nope" },
            },
        });

        await runWithLogContext(buildRequestLogContext({ id: "activity-9" }, {}), async () => {
            await assert.rejects(
                () => postLoops(
                    "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events",
                    body,
                    { headers: { "Content-Type": "application/json" } },
                    {
                        logger: capture.log,
                        post: async (url, data) => {
                            posted.push({ url, data });
                            throw backendError;
                        },
                    }
                ),
                (caught) => caught === backendError
            );
        });

        assert.equal(posted.length, 1);
        assert.equal(posted[0].data, body);

        const failed = capture.records().find((record) => record.event === "LOOPS_WEBHOOK_FAILED");
        assert.equal(failed.processingStatus, "webhook_failed");
        assert.equal(failed.status, 500);
        assert.equal(failed.class, "http_error");
        assert.equal(failed.dependency, "loops-webhook");
        assert.equal(failed.retryCount, 0);
        assert.equal(typeof failed.durationMs, "number");
        assert.equal(failed.error.name, "Error");
        assert.equal(failed.error.message, "backend down");
        assert.equal(failed.error.status, 500);
        assert.equal(failed.payload, body);
        assert.equal(failed.response, undefined);
        assert.equal(JSON.stringify(failed).includes("response-secret"), false);
        assert.equal(JSON.stringify(failed.error).includes("request-secret"), false);
    });
});

describe("telemetry failures", () => {
    it("does not reject when trackTrace throws", async () => {
        const capture = [];
        const loggerInstance = createLogger({
            level: "info",
            getClient: () => ({
                trackTrace: () => {
                    throw new Error("telemetry down");
                },
            }),
            writeStdout: (line) => capture.push(line),
            writeStderr: (line) => capture.push(line),
        });

        await assert.doesNotReject(async () => {
            loggerInstance.error("Loops webhook failed", {
                event: "LOOPS_WEBHOOK_FAILED",
                correlationId: "activity-42",
                response: { status: "error", message: "full backend body" },
            });
            await new Promise((resolve) => setImmediate(resolve));
        });

        const parsed = JSON.parse(capture.join("").trim().split("\n").pop());
        assert.equal(parsed.correlationId, "activity-42");
        assert.equal(parsed.response.message, "full backend body");
    });
});
