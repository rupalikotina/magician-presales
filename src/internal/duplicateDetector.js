// src/internal/duplicateDetector.js
// Process-local duplicate observation. The Map does not survive recycle and
// is not shared across App Service instances. A logged duplicate is telemetry;
// only the command path uses it to return early.
const { getLogStore, logDuplicate } = require("./logger");

const DEFAULT_TTL_MS = 5 * 60 * 1000;

function readIdentity(activity) {
    if (!activity || activity.id === undefined || activity.id === null || activity.id === "") {
        return undefined;
    }
    const tenant = activity.channelData && activity.channelData.tenant;
    const tenantId = tenant && tenant.id ? String(tenant.id) : undefined;
    const conversationId = activity.conversation && activity.conversation.id
        ? String(activity.conversation.id)
        : undefined;
    return {
        source: "teams",
        tenantId,
        conversationId,
        activityId: String(activity.id),
    };
}

function identityKey(identity) {
    return [
        identity.source,
        identity.tenantId || "",
        identity.conversationId || "",
        identity.activityId,
    ].join("\0");
}

/**
 * @param {{ ttlMs?: number, now?: () => number }} [options]
 */
function createDuplicateDetector(options = {}) {
    const ttlMs = options.ttlMs || DEFAULT_TTL_MS;
    const now = options.now || Date.now;
    const store = new Map();
    // One bucket per request context. A retry is a new context, so it counts again.
    const turns = new WeakMap();

    function prune(current) {
        for (const [key, entry] of store) {
            if (current - entry.firstSeen > ttlMs) {
                store.delete(key);
            }
        }
    }

    function turnBucket() {
        const ctx = getLogStore();
        if (!ctx) {
            return undefined;
        }
        let activities = turns.get(ctx);
        if (!activities) {
            activities = new Map();
            turns.set(ctx, activities);
        }
        return activities;
    }

    function observe(activity) {
        const identity = readIdentity(activity);
        if (!identity) {
            return { duplicate: false, skipped: true, sameTurn: false };
        }

        const current = now();
        prune(current);

        const activities = turnBucket();
        if (activities && activities.has(identity.activityId)) {
            return { ...activities.get(identity.activityId), sameTurn: true };
        }

        const key = identityKey(identity);
        let entry = store.get(key);
        if (entry && current - entry.firstSeen > ttlMs) {
            store.delete(key);
            entry = undefined;
        }

        let observation;
        if (entry) {
            entry.count += 1;
            entry.lastSeen = current;
            observation = {
                duplicate: true,
                skipped: false,
                sameTurn: false,
                ...identity,
                count: entry.count,
                firstSeen: entry.firstSeen,
                lastSeen: entry.lastSeen,
            };
        } else {
            const created = { count: 1, firstSeen: current, lastSeen: current };
            store.set(key, created);
            observation = {
                duplicate: false,
                skipped: false,
                sameTurn: false,
                ...identity,
                count: 1,
                firstSeen: created.firstSeen,
                lastSeen: created.lastSeen,
            };
        }

        if (activities) {
            activities.set(identity.activityId, observation);
        }
        return observation;
    }

    function reset() {
        store.clear();
    }

    return { observe, reset };
}

const detector = createDuplicateDetector();

/**
 * Record one sighting. Command callers pass suppress: true and return when
 * the result says so. Other callers ignore suppress and keep processing.
 * Tenant and conversation ids are logged as tenantId and conversationId so
 * the logger redactor can hash them. activityId stays raw.
 * @param {object} activity
 * @param {{ suppress?: boolean, text?: string, log?: Function }} [options]
 */
function noteActivity(activity, options = {}) {
    const observation = detector.observe(activity);
    const suppress = Boolean(options.suppress && observation.duplicate);
    if (observation.duplicate && !observation.sameTurn) {
        const emit = options.log || logDuplicate;
        emit(suppress ? "Suppressed duplicate handleCommandReceived" : "Duplicate Teams activity observed", {
            source: observation.source,
            tenantId: observation.tenantId,
            conversationId: observation.conversationId,
            activityId: observation.activityId,
            count: observation.count,
            firstSeen: observation.firstSeen,
            lastSeen: observation.lastSeen,
            activityType: activity && activity.type,
            text: options.text !== undefined ? options.text : activity && activity.text,
        });
    }
    return { ...observation, suppress };
}

module.exports = {
    DEFAULT_TTL_MS,
    createDuplicateDetector,
    noteActivity,
    resetDuplicateDetector: detector.reset,
};
