// src/internal/health.js
// Process health only. These handlers do not call Application Insights, Loops,
// or the Bot Connector, and they do not read credentials or connection strings.
// Restify requires async (req, res) or callback-style (req, res, next).
let listening = false;

async function live(_req, res) {
    res.send(200, { status: "live" });
}

async function ready(_req, res) {
    if (listening) {
        res.send(200, { status: "ready" });
        return;
    }
    res.send(503, { status: "starting" });
}

function setReady(value) {
    listening = Boolean(value);
}

module.exports = {
    live,
    ready,
    setReady,
};
