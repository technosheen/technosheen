import { JiraIssueSchema, type JiraIssue } from "@workday/contracts";
import type { JiraPort } from "./ports.js";
import { fetchJson } from "./http.js";

interface JiraSearchResponse {
  issues: Array<{
    key: string;
    fields: {
      summary: string;
      description?: unknown;
      status: { name: string };
      priority?: { name: string };
      duedate?: string | null;
      updated: string;
      assignee?: { displayName: string } | null;
      labels?: string[];
    };
  }>;
}

type JiraIssueResponse = JiraSearchResponse["issues"][number] & {
  fields: JiraSearchResponse["issues"][number]["fields"] & {
    comment?: {
      comments: Array<{
        id: string;
        author: { displayName: string };
        body: unknown;
        created: string;
      }>;
    };
  };
};

export interface JiraAdapterOptions {
  baseUrl: string;
  email: string;
  apiToken: string;
  fetcher?: typeof fetch;
}

export class JiraAdapter implements JiraPort {
  readonly #baseUrl: string;
  readonly #authorization: string;
  readonly #fetcher: typeof fetch;

  constructor(options: JiraAdapterOptions) {
    this.#baseUrl = options.baseUrl.replace(/\/$/, "");
    this.#authorization = `Basic ${Buffer.from(`${options.email}:${options.apiToken}`).toString("base64")}`;
    this.#fetcher = options.fetcher ?? fetch;
  }

  async listActiveAssignedIssues(): Promise<JiraIssue[]> {
    return this.#search(
      "assignee = currentUser() AND statusCategory != Done ORDER BY priority DESC, updated DESC"
    );
  }

  async listRecentlyUpdatedIssues(since: string): Promise<JiraIssue[]> {
    const jiraDate = since.slice(0, 16).replace("T", " ");
    return this.#search(`updated >= "${jiraDate}" ORDER BY updated DESC`);
  }

  async getIssueDetails(key: string) {
    const issue = await fetchJson<JiraIssueResponse>(
      `${this.#baseUrl}/rest/api/3/issue/${encodeURIComponent(key)}?fields=summary,description,status,priority,duedate,updated,assignee,labels,comment`,
      {
        headers: {
          authorization: this.#authorization,
          accept: "application/json"
        }
      },
      this.#fetcher
    );
    return {
      issue: mapIssue(issue),
      comments: (issue.fields.comment?.comments ?? []).map((comment) => ({
        author: comment.author.displayName,
        body: adfToText(comment.body),
        createdAt: new Date(comment.created).toISOString()
      }))
    };
  }

  async #search(jql: string): Promise<JiraIssue[]> {
    const response = await fetchJson<JiraSearchResponse>(
      `${this.#baseUrl}/rest/api/3/search/jql`,
      {
        method: "POST",
        headers: {
          authorization: this.#authorization,
          accept: "application/json",
          "content-type": "application/json"
        },
        body: JSON.stringify({
          jql,
          fields: ["summary", "description", "status", "priority", "duedate", "updated", "assignee", "labels"],
          maxResults: 100
        })
      },
      this.#fetcher
    );
    return response.issues.map(mapIssue);
  }
}

function mapIssue({ key, fields }: JiraSearchResponse["issues"][number]): JiraIssue {
  return JiraIssueSchema.parse({
    key,
    summary: fields.summary,
    description: adfToText(fields.description),
    status: fields.status.name,
    priority: normalizePriority(fields.priority?.name),
    dueDate: fields.duedate ?? null,
    updatedAt: new Date(fields.updated).toISOString(),
    assignee: fields.assignee?.displayName ?? null,
    mentionsCurrentUser: false,
    blocked: fields.status.name.toLowerCase().includes("block") || fields.labels?.includes("blocked")
  });
}

function adfToText(value: unknown): string {
  if (typeof value === "string") return value;
  if (!value || typeof value !== "object") return "";
  if ("text" in value && typeof value.text === "string") return value.text;
  if ("content" in value && Array.isArray(value.content)) {
    return value.content.map(adfToText).filter(Boolean).join(" ").replace(/\s+/g, " ").trim();
  }
  return "";
}

function normalizePriority(value: string | undefined): JiraIssue["priority"] {
  const normalized = value?.toLowerCase();
  if (normalized === "highest" || normalized === "critical") return "Highest";
  if (normalized === "high") return "High";
  if (normalized === "medium" || normalized === "major") return "Medium";
  if (normalized === "low" || normalized === "minor") return "Low";
  if (normalized === "lowest" || normalized === "trivial") return "Lowest";
  return "Unprioritized";
}
