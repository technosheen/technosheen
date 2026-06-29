import { describe, expect, it, vi } from "vitest";
import { buildApp } from "./app.js";
import { buildPlannerServices } from "./services.js";

describe("Fastify API", () => {
  it("returns the complete daily rundown contract", async () => {
    const services = buildPlannerServices({
      PORT: 4000,
      WEB_ORIGIN: "http://localhost:3000",
      PLANNER_USE_DEMO_DATA: true,
      PLANNER_TIME_ZONE: "America/New_York",
      PLANNER_WORKDAY_START: "09:00",
      PLANNER_WORKDAY_END: "17:00"
    });
    const app = await buildApp(services);
    const response = await app.inject({
      method: "POST",
      url: "/api/rundown/daily",
      payload: { date: "2026-06-29" }
    });

    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body).toMatchObject({
      date: "2026-06-29",
      summary: expect.any(String),
      priorities: expect.any(Array),
      blockers: expect.any(Array),
      actionItems: expect.any(Array),
      schedule: expect.any(Array),
      timesheet: { totalMinutes: 480 }
    });
  });

  it("passes only validated schedules to calendar sync", async () => {
    const syncCalendar = vi.fn().mockResolvedValue({
      created: 1,
      updated: 0,
      skippedMeetings: 1
    });
    const app = await buildApp({
      createRundown: vi.fn(),
      syncCalendar
    });
    const response = await app.inject({
      method: "POST",
      url: "/api/calendar/sync",
      payload: {
        schedule: [
          {
            kind: "work",
            id: "planner:HUB-235:1",
            issueKey: "HUB-235",
            title: "HUBSPOT | HUB-235",
            start: "2026-06-29T15:15:00.000Z",
            end: "2026-06-29T16:45:00.000Z",
            showAs: "free",
            plannerOwned: true,
            score: 90
          }
        ]
      }
    });

    expect(response.statusCode).toBe(200);
    expect(syncCalendar).toHaveBeenCalledOnce();
  });
});
