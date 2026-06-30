import type {
  DailyRundown,
  DailyRundownRequest,
  ScheduleItem
} from "@workday/contracts";
import {
  DemoJiraAdapter,
  ClientCredentialTokenProvider,
  DelegatedMicrosoftAuth,
  GraphClient,
  JiraAdapter,
  OutlookAdapter,
  TeamsAdapter,
  type JiraPort,
  type OutlookPort,
  type TeamsPort,
  type DeviceLogin,
  type DeviceLoginStatus,
  type MicrosoftAuthStatus
} from "@workday/integrations";
import {
  createDailyRundown,
  dateInTimeZone,
  zonedBoundary
} from "@workday/planner";
import type { AppConfig } from "./config.js";

export interface PlannerServices {
  createRundown(input?: DailyRundownRequest): Promise<DailyRundown>;
  syncCalendar(schedule: ScheduleItem[]): Promise<{
    created: number;
    updated: number;
    skippedMeetings: number;
  }>;
  microsoftAuth: {
    getStatus(): Promise<MicrosoftAuthStatus>;
    startDeviceLogin(): Promise<DeviceLogin>;
    getDeviceLoginStatus(sessionId: string): DeviceLoginStatus;
    disconnect(): Promise<void>;
  };
}

export function buildPlannerServices(config: AppConfig): PlannerServices {
  const jira = config.PLANNER_USE_DEMO_DATA
    ? new DemoJiraAdapter()
    : liveJiraAdapter(config);
  const microsoft = microsoftAdapters(config);

  return {
    async createRundown(input = { captures: { meetings: [], mail: [], teams: [] } }) {
      const date = input.date ?? dateInTimeZone(new Date(), config.PLANNER_TIME_ZONE);
      const dayStart = zonedBoundary(date, config.PLANNER_WORKDAY_START, config.PLANNER_TIME_ZONE);
      const dayEnd = zonedBoundary(date, config.PLANNER_WORKDAY_END, config.PLANNER_TIME_ZONE);
      const since = new Date(new Date(dayStart).getTime() - 24 * 60 * 60_000).toISOString();
      const [assignedIssues, recentlyUpdatedIssues, graphMeetings, graphMail, graphTeams] =
        await Promise.all([
          jira.listActiveAssignedIssues(),
          jira.listRecentlyUpdatedIssues(since),
          microsoft.outlook?.listMeetings(dayStart, dayEnd) ?? Promise.resolve([]),
          microsoft.outlook?.listUnreadMail(since) ?? Promise.resolve([]),
          microsoft.teams?.listRecentMessages(since) ?? Promise.resolve([])
      ]);
      const issues = [
        ...new Map(
          [...assignedIssues, ...recentlyUpdatedIssues].map((issue) => [issue.key, issue])
        ).values()
      ];
      return createDailyRundown({
        date,
        issues,
        meetings: mergeById(graphMeetings, input.captures.meetings),
        mail: mergeById(graphMail, input.captures.mail),
        teams: mergeById(graphTeams, input.captures.teams),
        dayStart,
        dayEnd
      });
    },
    async syncCalendar(schedule) {
      const blocks = schedule
        .filter((item) => item.kind === "work")
        .map(({ kind: _kind, ...block }) => block);
      if (!microsoft.outlook) {
        throw new Error("Calendar sync is disabled in explicit-capture mode.");
      }
      return microsoft.outlook.syncPlannerBlocks(blocks);
    },
    microsoftAuth: microsoft.auth
  };
}

type MicrosoftAdapters = {
  outlook?: OutlookPort;
  teams?: TeamsPort;
  auth: PlannerServices["microsoftAuth"];
};

function microsoftAdapters(config: AppConfig): MicrosoftAdapters {
  if (config.MICROSOFT_AUTH_MODE === "manual") {
    return {
      auth: {
        async getStatus() {
          return { mode: "manual", status: "connected", account: null };
        },
        async startDeviceLogin() {
          throw new Error("Microsoft sign-in is disabled in explicit-capture mode.");
        },
        getDeviceLoginStatus() {
          return { status: "failed", error: "Microsoft sign-in is disabled." };
        },
        async disconnect() {}
      }
    };
  }

  if (!config.MICROSOFT_CLIENT_ID) {
    throw new Error("Microsoft integration requires MICROSOFT_CLIENT_ID.");
  }
  const delegatedAuth =
    config.MICROSOFT_AUTH_MODE === "delegated"
      ? new DelegatedMicrosoftAuth({
          tenantId: config.MICROSOFT_TENANT_ID,
          clientId: config.MICROSOFT_CLIENT_ID,
          ...(config.MICROSOFT_TOKEN_CACHE_PATH
            ? { cachePath: config.MICROSOFT_TOKEN_CACHE_PATH }
            : {})
        })
      : null;
  if (config.MICROSOFT_AUTH_MODE === "client_credentials" && !config.MICROSOFT_CLIENT_SECRET) {
    throw new Error("Client-credentials auth requires MICROSOFT_CLIENT_SECRET.");
  }
  const tokenProvider =
    delegatedAuth ??
    new ClientCredentialTokenProvider({
      tenantId: config.MICROSOFT_TENANT_ID,
      clientId: config.MICROSOFT_CLIENT_ID,
      clientSecret: config.MICROSOFT_CLIENT_SECRET!
    });
  const graph = new GraphClient({ tokenProvider });
  const auth: PlannerServices["microsoftAuth"] = delegatedAuth
    ? delegatedAuth
    : {
        async getStatus() {
          return { mode: "client_credentials", status: "connected", account: null };
        },
        async startDeviceLogin() {
          throw new Error("Device sign-in is unavailable in client-credentials mode.");
        },
        getDeviceLoginStatus() {
          return { status: "failed", error: "Device sign-in is unavailable." };
        },
        async disconnect() {}
      };
  return {
    outlook: new OutlookAdapter({
      graph,
      ...(config.MICROSOFT_USER_ID ? { userId: config.MICROSOFT_USER_ID } : {})
    }),
    teams: new TeamsAdapter(graph, {
      ...(config.MICROSOFT_USER_ID ? { userId: config.MICROSOFT_USER_ID } : {}),
      includeChannels: config.MICROSOFT_ENABLE_TEAMS_CHANNELS
    }),
    auth
  };
}

function liveJiraAdapter(config: AppConfig): JiraPort {
  const required = [config.JIRA_BASE_URL, config.JIRA_EMAIL, config.JIRA_API_TOKEN];
  if (required.some((value) => !value)) {
    throw new Error("Live Jira integration requires Jira credentials.");
  }
  return new JiraAdapter({
    baseUrl: config.JIRA_BASE_URL!,
    email: config.JIRA_EMAIL!,
    apiToken: config.JIRA_API_TOKEN!
  });
}

function mergeById<T extends { id: string }>(source: T[], captured: T[]): T[] {
  return [...new Map([...source, ...captured].map((item) => [item.id, item])).values()];
}

export { dateInTimeZone, zonedBoundary };
