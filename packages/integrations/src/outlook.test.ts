import { describe, expect, it } from "vitest";
import type { WorkBlock } from "@workday/contracts";
import { OutlookAdapter } from "./outlook.js";

describe("Outlook free/busy sync", () => {
  it("creates free planner events while leaving real meetings untouched", async () => {
    const calls: Array<{ path: string; init?: RequestInit }> = [];
    const graph = {
      async request<T>(path: string, init?: RequestInit): Promise<T> {
        calls.push({ path, ...(init ? { init } : {}) });
        if (path.includes("calendarView")) {
          return {
            value: [
              {
                id: "real-meeting",
                subject: "Real meeting",
                body: { content: "<p>Human-owned</p>" },
                start: { dateTime: "2026-06-29T14:00:00Z", timeZone: "UTC" },
                end: { dateTime: "2026-06-29T15:00:00Z", timeZone: "UTC" },
                showAs: "busy"
              }
            ]
          } as T;
        }
        return {} as T;
      }
    };
    const adapter = new OutlookAdapter({ graph, userId: "sean@example.com" });
    const block: WorkBlock = {
      id: "planner:HUB-235:1",
      issueKey: "HUB-235",
      title: "HUBSPOT | HUB-235",
      start: "2026-06-29T15:15:00.000Z",
      end: "2026-06-29T16:45:00.000Z",
      showAs: "free",
      plannerOwned: true,
      score: 90
    };

    const result = await adapter.syncPlannerBlocks([block]);
    const createCall = calls.find(({ init }) => init?.method === "POST");
    const payload = JSON.parse(String(createCall?.init?.body)) as {
      showAs: string;
      body: { content: string };
    };

    expect(result).toEqual({ created: 1, updated: 0, skippedMeetings: 1 });
    expect(payload.showAs).toBe("free");
    expect(payload.body.content).toContain("workday-planner:planner:HUB-235:1");
    expect(calls.some(({ path, init }) => path.includes("real-meeting") && init)).toBe(false);
  });

  it("updates only an event carrying the planner marker", async () => {
    const methods: string[] = [];
    const graph = {
      async request<T>(path: string, init?: RequestInit): Promise<T> {
        if (init?.method) methods.push(`${init.method} ${path}`);
        if (path.includes("calendarView")) {
          return {
            value: [
              {
                id: "planner-event",
                subject: "Old",
                body: { content: "<!-- workday-planner:planner:HUB-235:1 -->" },
                start: { dateTime: "2026-06-29T15:15:00Z", timeZone: "UTC" },
                end: { dateTime: "2026-06-29T16:45:00Z", timeZone: "UTC" }
              }
            ]
          } as T;
        }
        return {} as T;
      }
    };
    const adapter = new OutlookAdapter({ graph, userId: "sean@example.com" });
    await adapter.syncPlannerBlocks([
      {
        id: "planner:HUB-235:1",
        issueKey: "HUB-235",
        title: "HUBSPOT | HUB-235",
        start: "2026-06-29T15:15:00.000Z",
        end: "2026-06-29T16:45:00.000Z",
        showAs: "free",
        plannerOwned: true,
        score: 90
      }
    ]);

    expect(methods).toEqual(["PATCH /users/sean%40example.com/events/planner-event"]);
  });
});
