import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import http from "node:http";
import { createInterface } from "node:readline";

const expected = {
  date: "2026-06-30",
  summary: "Test rundown",
  priorities: [],
  blockers: [],
  actionItems: [],
  schedule: [],
  timesheet: { entries: [], totalMinutes: 480 }
};

const api = http.createServer((request, response) => {
  assert.equal(request.url, "/api/rundown/daily");
  assert.equal(request.method, "POST");
  response.writeHead(200, { "content-type": "application/json" });
  response.end(JSON.stringify(expected));
});
api.listen(0, "127.0.0.1");
await once(api, "listening");

const address = api.address();
assert(address && typeof address === "object");

const server = spawn(process.execPath, ["./mcp/server.mjs"], {
  cwd: new URL("..", import.meta.url),
  env: {
    ...process.env,
    WORKDAY_PLANNER_API_URL: `http://127.0.0.1:${address.port}`
  },
  stdio: ["pipe", "pipe", "inherit"]
});
const lines = createInterface({ input: server.stdout });
const responses = [];
lines.on("line", (line) => responses.push(JSON.parse(line)));

server.stdin.write(
  `${JSON.stringify({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2025-03-26" }
  })}\n`
);
server.stdin.write(
  `${JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/list" })}\n`
);
server.stdin.write(
  `${JSON.stringify({
    jsonrpc: "2.0",
    id: 3,
    method: "tools/call",
    params: {
      name: "generate_workday_plan",
      arguments: { date: "2026-06-30" }
    }
  })}\n`
);

await waitFor(() => responses.length === 3);
assert.equal(responses[0].result.serverInfo.name, "cinch-workday-planner");
assert.equal(responses[1].result.tools[0].name, "generate_workday_plan");
assert.deepEqual(responses[2].result.structuredContent, expected);

server.kill();
api.close();
await once(api, "close");

async function waitFor(predicate) {
  const deadline = Date.now() + 5_000;
  while (!predicate()) {
    if (Date.now() > deadline) throw new Error("Timed out waiting for MCP responses");
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}
