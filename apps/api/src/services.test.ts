import { describe, expect, it } from "vitest";
import { zonedBoundary } from "./services.js";

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
});
