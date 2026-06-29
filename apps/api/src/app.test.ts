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
      PLANNER_WORKDAY_END: "17:00",
      MICROSOFT_AUTH_MODE: "manual",
      MICROSOFT_TENANT_ID: "organizations",
      MICROSOFT_ENABLE_TEAMS_CHANNELS: false
    });
    const app = await buildApp(services);
    const response = await app.inject({
      method: "POST",
      url: "/api/rundown/daily",
      payload: {
        date: "2026-06-29",
        captures: {
          meetings: [
            {
              id: "captured:meeting",
              title: "Captured planning session",
              start: "2026-06-29T14:00:00.000Z",
              end: "2026-06-29T14:30:00.000Z",
              showAs: "busy"
            }
          ],
          mail: [],
          teams: []
        }
      }
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
    expect(body.schedule).toContainEqual(
      expect.objectContaining({ id: "captured:meeting", kind: "meeting" })
    );
  });

  it("passes only validated schedules to calendar sync", async () => {
    const syncCalendar = vi.fn().mockResolvedValue({
      created: 1,
      updated: 0,
      skippedMeetings: 1
    });
    const app = await buildApp({
      createRundown: vi.fn(),
      syncCalendar,
      microsoftAuth: {
        getStatus: vi.fn(),
        startDeviceLogin: vi.fn(),
        getDeviceLoginStatus: vi.fn(),
        disconnect: vi.fn()
      }
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

  it("reports Microsoft connection state", async () => {
    const app = await buildApp(buildPlannerServices({
      PORT: 4000,
      WEB_ORIGIN: "http://localhost:3000",
      PLANNER_USE_DEMO_DATA: true,
      PLANNER_TIME_ZONE: "America/New_York",
      PLANNER_WORKDAY_START: "09:00",
      PLANNER_WORKDAY_END: "17:00",
      MICROSOFT_AUTH_MODE: "manual",
      MICROSOFT_TENANT_ID: "organizations",
      MICROSOFT_ENABLE_TEAMS_CHANNELS: false
    }));
    const response = await app.inject({
      method: "GET",
      url: "/api/auth/microsoft/status"
    });

    expect(response.statusCode).toBe(200);
    expect(response.json()).toEqual({
      mode: "manual",
      status: "connected",
      account: null
    });
  });

  it("protects automated daily triggers with a bearer token", async () => {
    const createRundown = vi.fn().mockResolvedValue({
      date: "2026-06-29",
      summary: "Ready",
      priorities: [],
      blockers: [],
      actionItems: [],
      schedule: [],
      timesheet: {
        entries: [{ id: "internal", code: "INT-58", description: "Internal", minutes: 480, source: "internal" }],
        totalMinutes: 480
      }
    });
    const app = await buildApp(
      {
        createRundown,
        syncCalendar: vi.fn(),
        microsoftAuth: {
          getStatus: vi.fn(),
          startDeviceLogin: vi.fn(),
          getDeviceLoginStatus: vi.fn(),
          disconnect: vi.fn()
        }
      },
      "http://localhost:3000",
      "automation-token-1234"
    );

    const unauthorized = await app.inject({
      method: "POST",
      url: "/api/automation/daily",
      payload: {}
    });
    const authorized = await app.inject({
      method: "POST",
      url: "/api/automation/daily",
      headers: { authorization: "Bearer automation-token-1234" },
      payload: {}
    });

    expect(unauthorized.statusCode).toBe(401);
    expect(authorized.statusCode).toBe(200);
    expect(createRundown).toHaveBeenCalledOnce();
  });
});
