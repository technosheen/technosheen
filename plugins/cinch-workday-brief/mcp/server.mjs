#!/usr/bin/env node

import { createInterface } from "node:readline";

const apiUrl = (process.env.WORKDAY_PLANNER_API_URL ?? "http://127.0.0.1:4000")
  .replace(/\/+$/, "");

const captureSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    date: {
      type: "string",
      description: "Planner date in YYYY-MM-DD format.",
      pattern: "^\\d{4}-\\d{2}-\\d{2}$"
    },
    meetings: {
      type: "array",
      description: "Meetings explicitly selected from the user's connected calendar.",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["id", "title", "start", "end"],
        properties: {
          id: { type: "string" },
          title: { type: "string" },
          start: { type: "string", format: "date-time" },
          end: { type: "string", format: "date-time" },
          showAs: { type: "string", enum: ["busy"] },
          webUrl: { type: "string", format: "uri" }
        }
      }
    },
    mail: {
      type: "array",
      description: "High-signal unread messages explicitly selected for planning.",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["id", "subject", "sender", "preview", "receivedAt"],
        properties: {
          id: { type: "string" },
          subject: { type: "string" },
          sender: { type: "string" },
          preview: { type: "string" },
          receivedAt: { type: "string", format: "date-time" },
          webUrl: { type: "string", format: "uri" }
        }
      }
    },
    teams: {
      type: "array",
      description: "Relevant Teams messages explicitly selected for planning.",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["id", "author", "content", "createdAt"],
        properties: {
          id: { type: "string" },
          author: { type: "string" },
          content: { type: "string" },
          createdAt: { type: "string", format: "date-time" },
          webUrl: { type: "string", format: "uri" }
        }
      }
    }
  }
};

const tools = [
  {
    name: "generate_workday_plan",
    description:
      "Generate a deterministic daily schedule and exactly eight-hour timesheet from context the user selected through approved Codex connectors. This tool does not read Microsoft Graph or write to any external system.",
    inputSchema: captureSchema
  }
];

const input = createInterface({ input: process.stdin, crlfDelay: Infinity });

input.on("line", async (line) => {
  if (!line.trim()) return;

  let request;
  try {
    request = JSON.parse(line);
  } catch {
    writeError(null, -32700, "Parse error");
    return;
  }

  if (request.method?.startsWith("notifications/")) return;

  try {
    switch (request.method) {
      case "initialize":
        writeResult(request.id, {
          protocolVersion: request.params?.protocolVersion ?? "2025-03-26",
          capabilities: { tools: {} },
          serverInfo: {
            name: "cinch-workday-planner",
            version: "0.2.0"
          },
          instructions:
            "Use generate_workday_plan only after the user or connector workflow has selected the relevant meetings, mail, and Teams messages. Never broaden retrieval, write to calendars, or claim the planner API fetched Microsoft data."
        });
        break;
      case "ping":
        writeResult(request.id, {});
        break;
      case "tools/list":
        writeResult(request.id, { tools });
        break;
      case "tools/call":
        await callTool(request);
        break;
      default:
        writeError(request.id ?? null, -32601, `Method not found: ${request.method}`);
    }
  } catch (error) {
    writeError(
      request.id ?? null,
      -32603,
      error instanceof Error ? error.message : "Unexpected planner tool error"
    );
  }
});

async function callTool(request) {
  if (request.params?.name !== "generate_workday_plan") {
    writeError(request.id, -32602, `Unknown tool: ${request.params?.name ?? ""}`);
    return;
  }

  const capture = request.params.arguments ?? {};
  const response = await fetch(`${apiUrl}/api/rundown/daily`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      ...(capture.date ? { date: capture.date } : {}),
      captures: {
        meetings: capture.meetings ?? [],
        mail: capture.mail ?? [],
        teams: capture.teams ?? []
      }
    })
  }).catch((error) => {
    throw new Error(
      `Workday Planner is unavailable at ${apiUrl}. Start it with npm run dev or set WORKDAY_PLANNER_API_URL. ${error instanceof Error ? error.message : ""}`.trim()
    );
  });

  const body = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(
      `Workday Planner returned HTTP ${response.status}: ${JSON.stringify(body)}`
    );
  }

  writeResult(request.id, {
    content: [
      {
        type: "text",
        text: JSON.stringify(body, null, 2)
      }
    ],
    structuredContent: body,
    isError: false
  });
}

function writeResult(id, result) {
  process.stdout.write(`${JSON.stringify({ jsonrpc: "2.0", id, result })}\n`);
}

function writeError(id, code, message) {
  process.stdout.write(
    `${JSON.stringify({ jsonrpc: "2.0", id, error: { code, message } })}\n`
  );
}
