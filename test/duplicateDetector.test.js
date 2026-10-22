// test/duplicateDetector.test.js
const { describe, it, beforeEach, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const axios = require("axios");
const { createDuplicateDetector, noteActivity, resetDuplicateDetector, DEFAULT_TTL_MS } = require("../src/internal/duplicateDetector");
const { runWithLogContext, toSchemaRecord } = require("../src/internal/logger");
const { hashIdentifier } = require("../src/internal/redact");
const { buildRequestLogContext } = require("../src/internal/requestContext");
const { LoopsCommandHandler } = require("../src/commands/loopsCommandHandler");

function activity(id, extras = {}) {
    return {
        id,
        type: "message",
        text: "hello from Teams",
        conversation: { id: "conversation-1" },
        channelData: { tenant: { id: "tenant-1" } },
        ...extras,
    };
}

describe("duplicate detector", () => {
    it("does not treat the first observation as a duplicate", () => {
        const detector = createDuplicateDetector({ now: () => 1_000 });
        const first = detector.observe(activity("activity-1"));
        assert.equal(first.duplicate, false);
        assert.equal(first.skipped, false);
        assert.equal(first.count, 1);
        assert.equal(first.firstSeen, 1_000);
        assert.equal(first.lastSeen, 1_000);
        assert.equal(first.source, "teams");
        assert.equal(first.activityId, "activity-1");
    });

    it("returns a duplicate within the TTL with count 2 and the original firstSeen", () => {
        let now = 5_000;
        const detector = createDuplicateDetector({ now: () => now });
        const first = detector.observe(activity("activity-1"));
        now += 1_000;
        const second = detector.observe(activity("activity-1"));
        assert.equal(second.duplicate, true);
        assert.equal(second.count, 2);
        assert.equal(second.firstSeen, first.firstSeen);
        assert.equal(second.lastSeen, now);
    });

    it("treats an observation after the TTL as new", () => {
        let now = 10_000;
        const detector = createDuplicateDetector({ now: () => now, ttlMs: DEFAULT_TTL_MS });
        detector.observe(activity("activity-1"));
        now += DEFAULT_TTL_MS + 1;
        const again = detector.observe(activity("activity-1"));
        assert.equal(again.duplicate, false);
        assert.equal(again.count, 1);
        assert.equal(again.firstSeen, now);
    });

    it("does not treat a different activity id as a duplicate", () => {
        const detector = createDuplicateDetector({ now: () => 2_000 });
        detector.observe(activity("activity-1"));
        const other = detector.observe(activity("activity-2"));
        assert.equal(other.duplicate, false);
        assert.equal(other.count, 1);
        assert.equal(other.activityId, "activity-2");
    });

    it("skips detection when the activity id is missing", () => {
        const detector = createDuplicateDetector({ now: () => 3_000 });
        const first = detector.observe({ conversation: { id: "conversation-1" }, text: "one" });
        const second = detector.observe({ conversation: { id: "conversation-2" }, text: "two" });
        assert.equal(first.skipped, true);
        assert.equal(first.duplicate, false);
        assert.equal(second.skipped, true);
        assert.equal(second.duplicate, false);
    });

    it("counts a repeat once per turn when message and command handlers both see it", async () => {
        let now = 8_000;
        const detector = createDuplicateDetector({ now: () => now });
        const body = activity("activity-7");
        await runWithLogContext(buildRequestLogContext(body, {}), async () => {
            const message = detector.observe(body);
            const command = detector.observe(body);
            assert.equal(message.duplicate, false);
            assert.equal(command.sameTurn, true);
            assert.equal(command.duplicate, false);
            assert.equal(command.count, 1);
        });
        now += 50;
        await runWithLogContext(buildRequestLogContext(body, {}), async () => {
            const retry = detector.observe(body);
            assert.equal(retry.duplicate, true);
            assert.equal(retry.count, 2);
            assert.equal(retry.firstSeen, 8_000);
        });
    });

    it("hashes tenant and conversation ids and leaves the activity id raw", () => {
        const previous = process.env.LOG_RAW_IDENTIFIERS;
        process.env.LOG_RAW_IDENTIFIERS = "false";
        const captured = [];
        try {
            resetDuplicateDetector();
            const body = activity("activity-raw");
            noteActivity(body, { suppress: false, log: () => {} });
            noteActivity(body, {
                suppress: false,
                log: (_message, fields) => captured.push(fields),
            });
            const record = toSchemaRecord({
                level: "warn",
                message: "Duplicate Teams activity observed",
                timestamp: "2026-01-01T00:00:00.000Z",
                event: "TEAMS_DUPLICATE_ACTIVITY_DETECTED",
                ...captured[0],
            });
            assert.equal(record.activityId, "activity-raw");
            assert.equal(record.tenantId, hashIdentifier("tenant-1"));
            assert.equal(record.conversationId, hashIdentifier("conversation-1"));
            assert.equal(record.event, "TEAMS_DUPLICATE_ACTIVITY_DETECTED");
            assert.equal(record.count, 2);
        } finally {
            resetDuplicateDetector();
            if (previous === undefined) {
                delete process.env.LOG_RAW_IDENTIFIERS;
            } else {
                process.env.LOG_RAW_IDENTIFIERS = previous;
            }
        }
    });
});

describe("duplicate handling", () => {
    let originalPost;

    beforeEach(() => {
        resetDuplicateDetector();
        originalPost = axios.post;
    });

    afterEach(() => {
        axios.post = originalPost;
        resetDuplicateDetector();
    });

    it("keeps processing a repeated message and returns early for a repeated command", async () => {
        const body = activity("activity-command", { text: "loops status" });
        const messageObservation = noteActivity(body, { suppress: false, log: () => {} });
        assert.equal(messageObservation.suppress, false);

        const repeatedMessage = noteActivity(body, { suppress: false, log: () => {} });
        assert.equal(repeatedMessage.duplicate, true);
        assert.equal(repeatedMessage.suppress, false);
        assert.equal(repeatedMessage.count, 2);

        let posts = 0;
        axios.post = async () => {
            posts += 1;
            throw new Error("duplicate command was forwarded");
        };
        const handler = new LoopsCommandHandler();
        const result = await handler.handleCommandReceived(
            { activity: body },
            { text: body.text }
        );
        assert.equal(result, undefined);
        assert.equal(posts, 0);
    });

    it("does not suppress a command that was already observed in the same turn", async () => {
        const body = activity("activity-same-turn", { text: "loops status" });
        let posts = 0;
        axios.post = async (url, data) => {
            posts += 1;
            assert.equal(url, "https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events");
            assert.equal(typeof data, "string");
            return { status: 200, data: { status: "ok", cardJson: "{\"type\":\"AdaptiveCard\",\"version\":\"1.4\",\"body\":[]}" } };
        };

        await runWithLogContext(buildRequestLogContext(body, {}), async () => {
            const observed = noteActivity(body, { suppress: false, log: () => {} });
            assert.equal(observed.duplicate, false);
            assert.equal(observed.suppress, false);
            const handler = new LoopsCommandHandler();
            await handler.handleCommandReceived(
                {
                    activity: body,
                    sendActivity: async () => ({ id: "sent" }),
                    updateActivity: async () => ({ id: "updated" }),
                },
                { text: body.text, matches: ["loops status", "loops", " status", "status"] }
            );
        });
        assert.ok(posts >= 1);
    });
});
