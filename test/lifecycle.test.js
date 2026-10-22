// test/lifecycle.test.js
const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { EventEmitter } = require("events");
const { createLifecycle, FLUSH_TIMEOUT_MS, flushTelemetry } = require("../src/internal/lifecycle");

function harness(overrides = {}) {
    const order = [];
    const client = overrides.client || {
        flush(options) {
            order.push("flush");
            if (options && typeof options.callback === "function" && !overrides.hang) {
                options.callback("flushed");
            }
        },
    };
    const lifecycle = createLifecycle({
        timeoutMs: overrides.timeoutMs,
        exit: (code) => {
            order.push(`exit:${code}`);
        },
        getClient: () => client,
        logInfo: (_message, fields) => {
            order.push(fields.event);
        },
        logError: (_message, fields) => {
            order.push(fields.event);
            if (overrides.capture) {
                overrides.capture.push(fields);
            }
        },
        ...overrides.deps,
    });
    return { order, client, lifecycle };
}

describe("process lifecycle", () => {
    it("logs BOT_STARTED with the host identity and start time", () => {
        const fields = [];
        const { lifecycle } = harness({
            deps: {
                logInfo: (_message, record) => fields.push(record),
            },
        });
        lifecycle.onStarted({ port: 3978, botAppIdSuffix: "abcd1234" });
        assert.equal(fields[0].event, "BOT_STARTED");
        assert.equal(fields[0].hosting, "azure-app-service-windows");
        assert.equal(typeof fields[0].pid, "number");
        assert.equal(typeof fields[0].nodeVersion, "string");
        assert.equal(typeof fields[0].version, "string");
        assert.equal(typeof fields[0].instanceId, "string");
        assert.equal(typeof fields[0].startTime, "string");
        assert.equal(fields[0].port, 3978);
    });

    it("logs BOT_STARTUP_FAILED before the process exits", () => {
        const source = fs.readFileSync(path.join(__dirname, "..", "src", "index.js"), "utf8");
        const failedAt = source.indexOf("onStartupFailed(");
        const exitAt = source.indexOf("process.exit(1)", failedAt);
        assert.ok(failedAt > 0);
        assert.ok(exitAt > failedAt);
        const events = [];
        const { lifecycle } = harness({ capture: events });
        lifecycle.onStartupFailed("Set MicrosoftAppId/MicrosoftAppPassword (or BOT_ID/SECRET_BOT_PASSWORD/BOT_PASSWORD).");
        assert.equal(events[0].event, "BOT_STARTUP_FAILED");
    });

    it("flushes telemetry and then exits on shutdown", async () => {
        const { order, lifecycle } = harness();
        await lifecycle.shutdown("SIGTERM");
        assert.deepEqual(order, ["BOT_STOPPING", "flush", "exit:0"]);
    });

    it("abandons a hung flush at the timeout and still exits", async () => {
        assert.equal(FLUSH_TIMEOUT_MS, 2000);
        const { order, lifecycle } = harness({ hang: true, timeoutMs: 40 });
        const started = Date.now();
        await lifecycle.shutdown("SIGINT");
        const elapsed = Date.now() - started;
        assert.deepEqual(order, ["BOT_STOPPING", "flush", "exit:0"]);
        assert.ok(elapsed >= 30);
        assert.ok(elapsed < 1000);
    });

    it("exits 1 after an uncaught exception is logged and flushed", async () => {
        const tracked = [];
        const { order, lifecycle } = harness({
            client: {
                trackException(payload) {
                    tracked.push(payload.exception.message);
                },
                flush(options) {
                    order.push("flush");
                    options.callback();
                },
            },
        });
        await lifecycle.onUncaughtException(new Error("crash"));
        assert.deepEqual(order, ["BOT_UNHANDLED_EXCEPTION", "flush", "exit:1"]);
        assert.deepEqual(tracked, ["crash"]);
    });

    it("logs an unhandled rejection and does not exit", () => {
        const captured = [];
        const { order, lifecycle } = harness({ capture: captured });
        lifecycle.onUnhandledRejection(new Error("rejected"));
        assert.deepEqual(order, ["BOT_UNHANDLED_REJECTION"]);
        assert.equal(captured[0].error.message, "rejected");
    });

    it("registers signal and crash handlers without exiting on rejection", async () => {
        const emitter = new EventEmitter();
        const { order, lifecycle } = harness();
        lifecycle.register(emitter);
        emitter.emit("unhandledRejection", new Error("still up"));
        assert.deepEqual(order, ["BOT_UNHANDLED_REJECTION"]);
        emitter.emit("SIGTERM");
        await new Promise((resolve) => setImmediate(resolve));
        assert.deepEqual(order, ["BOT_UNHANDLED_REJECTION", "BOT_STOPPING", "flush", "exit:0"]);
    });
});

describe("flushTelemetry", () => {
    it("resolves immediately when the client has no flush method", async () => {
        const started = Date.now();
        await flushTelemetry(null, { timeoutMs: 500 });
        assert.ok(Date.now() - started < 100);
    });
});
