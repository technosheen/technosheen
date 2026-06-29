import { mkdir } from "node:fs/promises";
import { dirname, join } from "node:path";
import {
  PublicClientApplication,
  type AccountInfo,
  type IPublicClientApplication
} from "@azure/msal-node";

export const DEFAULT_DELEGATED_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Calendars.ReadWrite",
  "Mail.Read",
  "Chat.Read"
] as const;

export interface AccessTokenProvider {
  getAccessToken(): Promise<string>;
}

export interface MicrosoftAuthStatus {
  mode: "manual" | "delegated" | "client_credentials" | "demo";
  status: "connected" | "disconnected";
  account: { name: string | null; username: string } | null;
}

export interface DeviceLogin {
  sessionId: string;
  userCode: string;
  verificationUri: string;
  message: string;
  expiresIn: number;
}

export interface DeviceLoginStatus {
  status: "pending" | "connected" | "failed";
  account?: { name: string | null; username: string };
  error?: string;
}

export interface DelegatedMicrosoftAuthOptions {
  clientId: string;
  tenantId?: string;
  cachePath?: string;
  scopes?: string[];
  application?: IPublicClientApplication;
}

interface LoginSession {
  status: DeviceLoginStatus;
}

export class MicrosoftAuthRequiredError extends Error {
  constructor() {
    super("Connect Microsoft 365 before using live Outlook or Teams integrations.");
    this.name = "MicrosoftAuthRequiredError";
  }
}

export class DelegatedMicrosoftAuth implements AccessTokenProvider {
  readonly #scopes: string[];
  readonly #application: Promise<IPublicClientApplication>;
  readonly #sessions = new Map<string, LoginSession>();

  constructor(options: DelegatedMicrosoftAuthOptions) {
    this.#scopes = options.scopes ?? [...DEFAULT_DELEGATED_SCOPES];
    this.#application = options.application
      ? Promise.resolve(options.application)
      : createPublicClientApplication(options);
  }

  async getStatus(): Promise<MicrosoftAuthStatus> {
    const account = await this.#getAccount();
    return {
      mode: "delegated",
      status: account ? "connected" : "disconnected",
      account: account ? accountSummary(account) : null
    };
  }

  async startDeviceLogin(): Promise<DeviceLogin> {
    const application = await this.#application;
    const sessionId = crypto.randomUUID();
    const session: LoginSession = { status: { status: "pending" } };
    this.#sessions.set(sessionId, session);

    let resolvePrompt: (value: DeviceLogin) => void;
    let rejectPrompt: (reason: unknown) => void;
    const prompt = new Promise<DeviceLogin>((resolve, reject) => {
      resolvePrompt = resolve;
      rejectPrompt = reject;
    });

    void application
      .acquireTokenByDeviceCode({
        scopes: this.#scopes,
        deviceCodeCallback: (response) => {
          resolvePrompt({
            sessionId,
            userCode: response.userCode,
            verificationUri: response.verificationUri,
            message: response.message,
            expiresIn: response.expiresIn
          });
        }
      })
      .then((result) => {
        if (!result?.account) {
          session.status = { status: "failed", error: "Microsoft sign-in returned no account." };
          return;
        }
        session.status = { status: "connected", account: accountSummary(result.account) };
      })
      .catch((error: unknown) => {
        const message = error instanceof Error ? error.message : "Microsoft sign-in failed.";
        session.status = { status: "failed", error: message };
        rejectPrompt(error);
      });

    return prompt;
  }

  getDeviceLoginStatus(sessionId: string): DeviceLoginStatus {
    return this.#sessions.get(sessionId)?.status ?? {
      status: "failed",
      error: "Unknown or expired sign-in session."
    };
  }

  async disconnect(): Promise<void> {
    const application = await this.#application;
    for (const account of await application.getAllAccounts()) {
      await application.getTokenCache().removeAccount(account);
    }
  }

  async getAccessToken(): Promise<string> {
    const application = await this.#application;
    const account = (await application.getAllAccounts())[0];
    if (!account) throw new MicrosoftAuthRequiredError();
    const result = await application.acquireTokenSilent({
      account,
      scopes: this.#scopes
    });
    return result.accessToken;
  }

  async #getAccount(): Promise<AccountInfo | undefined> {
    return (await (await this.#application).getAllAccounts())[0];
  }
}

async function createPublicClientApplication(
  options: DelegatedMicrosoftAuthOptions
): Promise<IPublicClientApplication> {
  const {
    DataProtectionScope,
    Environment,
    PersistenceCachePlugin,
    PersistenceCreator
  } = await import("@azure/msal-node-extensions");
  const home = Environment.getUserRootDirectory();
  if (!home && !options.cachePath) throw new Error("Unable to locate a secure token-cache path.");
  const cachePath =
    options.cachePath ?? join(home!, ".workday-planner", "microsoft-token-cache.json");
  await mkdir(dirname(cachePath), { recursive: true, mode: 0o700 });
  const persistence = await PersistenceCreator.createPersistence({
    cachePath,
    dataProtectionScope: DataProtectionScope.CurrentUser,
    serviceName: "workday-planner",
    accountName: options.clientId,
    usePlaintextFileOnLinux: false
  });
  return new PublicClientApplication({
    auth: {
      clientId: options.clientId,
      authority: `https://login.microsoftonline.com/${options.tenantId ?? "organizations"}`
    },
    cache: { cachePlugin: new PersistenceCachePlugin(persistence) }
  });
}

function accountSummary(account: AccountInfo) {
  return {
    name: account.name ?? null,
    username: account.username
  };
}
