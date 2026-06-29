import type { IPublicClientApplication } from "@azure/msal-node";
import { describe, expect, it, vi } from "vitest";
import { DelegatedMicrosoftAuth, MicrosoftAuthRequiredError } from "./microsoft-auth.js";

describe("DelegatedMicrosoftAuth", () => {
  it("requires an interactive connection when no account is cached", async () => {
    const application = fakeApplication();
    const auth = new DelegatedMicrosoftAuth({
      clientId: "client-id",
      application
    });

    await expect(auth.getStatus()).resolves.toMatchObject({ status: "disconnected" });
    await expect(auth.getAccessToken()).rejects.toBeInstanceOf(MicrosoftAuthRequiredError);
  });

  it("returns a silent token for the cached signed-in account", async () => {
    const account = {
      homeAccountId: "home",
      environment: "login.microsoftonline.com",
      tenantId: "tenant",
      username: "sean@example.com",
      localAccountId: "local",
      name: "Sean"
    };
    const application = fakeApplication({
      getAllAccounts: vi.fn().mockResolvedValue([account]),
      acquireTokenSilent: vi.fn().mockResolvedValue({ accessToken: "graph-token" })
    });
    const auth = new DelegatedMicrosoftAuth({
      clientId: "client-id",
      application
    });

    await expect(auth.getAccessToken()).resolves.toBe("graph-token");
    await expect(auth.getStatus()).resolves.toEqual({
      mode: "delegated",
      status: "connected",
      account: { name: "Sean", username: "sean@example.com" }
    });
  });
});

function fakeApplication(overrides: Record<string, unknown> = {}): IPublicClientApplication {
  return {
    getAllAccounts: vi.fn().mockResolvedValue([]),
    acquireTokenSilent: vi.fn(),
    acquireTokenByDeviceCode: vi.fn(),
    getTokenCache: vi.fn().mockReturnValue({ removeAccount: vi.fn() }),
    ...overrides
  } as unknown as IPublicClientApplication;
}
