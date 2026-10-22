// test/health.test.js
const { describe, it, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { live, ready, setReady } = require("../src/internal/health");
const { postLoops } = require("../src/internal/loopsClient");
const { createLogger, runWithLogContext } = require("../src/internal/logger");
const { buildRequestLogContext } = require("../src/internal/requestContext");

function fakeRes() {
    return {
        code: null,
        body: null,
        send(status, body) {
            this.code = status;
            this.body = body;
        },
    };
}

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

describe("health routes", () => {
    afterEach(() => {
        setReady(false);
    });

    it("live returns 200 and does not read environment secrets", () => {
        const source = fs.readFileSync(path.join(__dirname, "..", "src", "internal", "health.js"), "utf8");
        const index = fs.readFileSync(path.join(__dirname, "..", "src", "index.js"), "utf8");
        assert.equal(source.includes("process.env"), false);
        assert.match(index, /server\.get\("\/health\/live", live\)/);
        assert.match(index, /server\.get\("\/health\/ready", ready\)/);
        assert.match(index, /setReady\(true\)/);

        const previousPassword = process.env.MicrosoftAppPassword;
        const previousConnection = process.env.APPLICATIONINSIGHTS_CONNECTION_STRING;
        process.env.MicrosoftAppPassword = "super-secret-password";
        process.env.APPLICATIONINSIGHTS_CONNECTION_STRING = "InstrumentationKey=secret-key";
        try {
            const res = fakeRes();
            live({}, res);
            assert.equal(res.code, 200);
            assert.deepEqual(res.body, { status: "live" });
            const encoded = JSON.stringify(res.body);
            assert.equal(encoded.includes("super-secret-password"), false);
            assert.equal(encoded.includes("secret-key"), false);
        } finally {
            if (previousPassword === undefined) {
                delete process.env.MicrosoftAppPassword;
            } else {
                process.env.MicrosoftAppPassword = previousPassword;
            }
            if (previousConnection === undefined) {
                delete process.env.APPLICATIONINSIGHTS_CONNECTION_STRING;
            } else {
                process.env.APPLICATIONINSIGHTS_CONNECTION_STRING = previousConnection;
            }
        }
    });

    it("ready is 503 before listen and 200 after it", () => {
        setReady(false);
        const starting = fakeRes();
        ready({}, starting);
        assert.equal(starting.code, 503);
        assert.deepEqual(starting.body, { status: "starting" });

        setReady(true);
        const listening = fakeRes();
        ready({}, listening);
        assert.equal(listening.code, 200);
        assert.deepEqual(listening.body, { status: "ready" });
    });
});

describe("Loops dependency classes", () => {
    const target = "api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events";
    const url = `https://${target}?token=query-secret`;

    async function failWith(error) {
        const capture = captureLogger();
        await runWithLogContext(buildRequestLogContext({ id: "activity-7" }, {}), async () => {
            await assert.rejects(
                () => postLoops(url, JSON.stringify({ messageText: "hello" }), {}, {
                    logger: capture.log,
                    post: async () => {
                        throw error;
                    },
                }),
                (caught) => caught === error
            );
        });
        return capture.records().find((record) => record.event === "LOOPS_WEBHOOK_FAILED");
    }

    function assertDependency(record, errorClass, status) {
        assert.equal(record.dependency, "loops-webhook");
        assert.equal(record.method, "POST");
        assert.equal(record.target, target);
        assert.equal(record.class, errorClass);
        assert.equal(record.retryCount, 0);
        assert.equal(record.correlationId, "activity-7");
        assert.equal(typeof record.durationMs, "number");
        if (status) {
            assert.equal(record.status, status);
        }
        const encoded = JSON.stringify(record);
        assert.equal(encoded.includes("response-secret"), false);
        assert.equal(encoded.includes("query-secret"), false);
        assert.equal(record.response, undefined);
    }

    it("maps a timeout without the response body", async () => {
        const error = Object.assign(new Error("timed out"), {
            code: "ECONNABORTED",
            response: { status: 500, data: { token: "response-secret" } },
        });
        const record = await failWith(error);
        assertDependency(record, "timeout", 500);

        const timedOut = Object.assign(new Error("timed out"), { code: "ETIMEDOUT" });
        const second = await failWith(timedOut);
        assertDependency(second, "timeout");
    });

    it("maps HTTP 500 without the response body", async () => {
        const error = Object.assign(new Error("unavailable"), {
            response: { status: 500, data: { token: "response-secret", detail: "nope" } },
        });
        const record = await failWith(error);
        assertDependency(record, "http_error", 500);
    });

    it("maps a network error without the response body", async () => {
        const error = Object.assign(new Error("connection refused"), {
            code: "ECONNREFUSED",
            response: { data: { token: "response-secret" } },
        });
        const record = await failWith(error);
        assertDependency(record, "network_error");
    });

    it("keeps Teams connector failures off the Loops dependency name", () => {
        const source = fs.readFileSync(path.join(__dirname, "..", "src", "internal", "teamsResponse.js"), "utf8");
        assert.equal(source.includes("loops-webhook"), false);
        assert.match(source, /TEAMS_RESPONSE_FAILED/);
        assert.match(source, /TEAMS_AUTH_FAILED/);
    });
});
