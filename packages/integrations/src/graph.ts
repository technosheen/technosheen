import { fetchJson } from "./http.js";

interface TokenResponse {
  access_token: string;
  expires_in: number;
}

export interface GraphClientOptions {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  fetcher?: typeof fetch;
}

export class GraphClient {
  readonly #options: GraphClientOptions;
  readonly #fetcher: typeof fetch;
  #token: { value: string; expiresAt: number } | undefined;

  constructor(options: GraphClientOptions) {
    this.#options = options;
    this.#fetcher = options.fetcher ?? fetch;
  }

  async request<T>(path: string, init: RequestInit = {}): Promise<T> {
    const token = await this.#getToken();
    return fetchJson<T>(
      `https://graph.microsoft.com/v1.0${path}`,
      {
        ...init,
        headers: {
          authorization: `Bearer ${token}`,
          accept: "application/json",
          ...(init.body ? { "content-type": "application/json" } : {}),
          ...init.headers
        }
      },
      this.#fetcher
    );
  }

  async #getToken(): Promise<string> {
    if (this.#token && this.#token.expiresAt > Date.now() + 60_000) return this.#token.value;
    const body = new URLSearchParams({
      client_id: this.#options.clientId,
      client_secret: this.#options.clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    });
    const response = await fetchJson<TokenResponse>(
      `https://login.microsoftonline.com/${this.#options.tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body
      },
      this.#fetcher
    );
    this.#token = {
      value: response.access_token,
      expiresAt: Date.now() + response.expires_in * 1000
    };
    return response.access_token;
  }
}
