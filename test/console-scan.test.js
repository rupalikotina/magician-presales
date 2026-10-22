// test/console-scan.test.js
const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");

const CONSOLE_CALL = /console\.(log|warn|error|debug)\s*\(/;

function listJsFiles(dir, found = []) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const fullPath = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            listJsFiles(fullPath, found);
        } else if (entry.name.endsWith(".js")) {
            found.push(fullPath);
        }
    }
    return found;
}

function readSrc(relativePath) {
    return fs.readFileSync(path.join(__dirname, "..", relativePath), "utf8");
}

describe("application console calls", () => {
    it("src/ has no console.log, console.warn, console.error, or console.debug calls", () => {
        const srcDir = path.join(__dirname, "..", "src");
        const hits = [];
        for (const file of listJsFiles(srcDir)) {
            const lines = fs.readFileSync(file, "utf8").split("\n");
            lines.forEach((line, index) => {
                if (CONSOLE_CALL.test(line)) {
                    hits.push(`${path.relative(path.join(__dirname, ".."), file)}:${index + 1}: ${line.trim()}`);
                }
            });
        }
        assert.deepEqual(hits, []);
    });

    it("keeps Teams text, payloads, request bodies, backend responses, and action data", () => {
        const teamsBot = readSrc("src/teamsBot.js");
        const index = readSrc("src/index.js");
        const commands = readSrc("src/commands/loopsCommandHandler.js");
        const loopsClient = readSrc("src/internal/loopsClient.js");

        assert.match(teamsBot, /logger\.info\("Received message",\s*\{[\s\S]*?\btext,/);
        assert.match(teamsBot, /messageText:\s*context\.activity\.text/);
        assert.match(teamsBot, /logger\.info\("Backend response",\s*\{[\s\S]*?\bresponse:\s*response\.data/);
        assert.match(teamsBot, /logger\.info\("Received submit action data",\s*\{[\s\S]*?\bactionData,/);
        assert.match(teamsBot, /rawEvent:\s*context/);

        assert.match(index, /logger\.info\("Request received",\s*\{[\s\S]*?\bbody:\s*req\.body/);
        assert.match(index, /text:\s*activity\.text/);

        assert.match(commands, /logger\.info\("Backend response",\s*\{[\s\S]*?\bresponse:\s*response\.data/);
        assert.match(commands, /text:\s*message\.text/);

        assert.match(loopsClient, /payload:\s*data/);
        assert.match(loopsClient, /response:\s*responseBody\(response\)/);
    });
});
