import { describe, expect, it } from "vitest";
import { readConfig } from "./config.js";
import { dateInTimeZone, zonedBoundary } from "./services.js";

describe("zonedBoundary", () => {
  it("converts Eastern daylight time to UTC", () => {
    expect(zonedBoundary("2026-06-29", "09:00", "America/New_York")).toBe(
      "2026-06-29T13:00:00.000Z"
    );
  });

  it("accounts for Eastern standard time", () => {
    expect(zonedBoundary("2026-12-29", "09:00", "America/New_York")).toBe(
      "2026-12-29T14:00:00.000Z"
    );
  });

  it("treats blank optional credentials as unconfigured", () => {
    const config = readConfig({
      JIRA_API_TOKEN: "",
      MICROSOFT_CLIENT_ID: "",
      PLANNER_AUTOMATION_TOKEN: ""
    });

    expect(config.JIRA_API_TOKEN).toBeUndefined();
    expect(config.MICROSOFT_CLIENT_ID).toBeUndefined();
    expect(config.PLANNER_AUTOMATION_TOKEN).toBeUndefined();
    expect(config.MICROSOFT_AUTH_MODE).toBe("manual");
  });

  it("uses the configured timezone when deriving a work date", () => {
    const instant = new Date("2026-06-30T02:00:00.000Z");
    expect(dateInTimeZone(instant, "America/New_York")).toBe("2026-06-29");
    expect(dateInTimeZone(instant, "Asia/Tokyo")).toBe("2026-06-30");
  });
});
