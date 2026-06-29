import type {
  CalendarSyncResponse,
  JiraIssue,
  MailMessage,
  Meeting,
  TeamsMessage,
  WorkBlock
} from "@workday/contracts";
import type { JiraPort, OutlookPort, TeamsPort } from "./ports.js";

export class DemoJiraAdapter implements JiraPort {
  async listActiveAssignedIssues(): Promise<JiraIssue[]> {
    return [
      {
        key: "HUB-235",
        summary: "Shared fragments in HubSpot",
        description: "Review global fragments and test cases.",
        status: "Ready for QA",
        priority: "High",
        dueDate: null,
        updatedAt: "2026-06-29T12:05:00.000Z",
        assignee: "Sean Mahoney",
        mentionsCurrentUser: true,
        blocked: false
      },
      {
        key: "HUB-62",
        summary: "Interactive USA map module",
        description: "Complete responsive map implementation.",
        status: "Dev In Progress",
        priority: "High",
        dueDate: null,
        updatedAt: "2026-06-26T15:12:00.000Z",
        assignee: "Sean Mahoney",
        mentionsCurrentUser: false,
        blocked: false
      },
      {
        key: "DTC-5238",
        summary: "Boost Coverage component",
        description: "Clarify acceptance criteria and Figma alignment.",
        status: "Blocked",
        priority: "Medium",
        dueDate: "2026-06-29",
        updatedAt: "2026-06-29T12:08:00.000Z",
        assignee: "Damen Golden",
        mentionsCurrentUser: false,
        blocked: true
      }
    ];
  }

  async listRecentlyUpdatedIssues(): Promise<JiraIssue[]> {
    return this.listActiveAssignedIssues();
  }

  async getIssueDetails(key: string) {
    const issues = await this.listActiveAssignedIssues();
    const issue = issues.find((candidate) => candidate.key === key);
    if (!issue) throw new Error(`Demo issue ${key} not found`);
    return { issue, comments: [] };
  }
}

export class DemoOutlookAdapter implements OutlookPort {
  readonly synced: WorkBlock[] = [];

  async listMeetings(start: string): Promise<Meeting[]> {
    const date = start.slice(0, 10);
    return [
      {
        id: "meeting:standup",
        title: "HubSpot Stand Up",
        start: `${date}T13:30:00.000Z`,
        end: `${date}T14:00:00.000Z`,
        showAs: "busy"
      },
      {
        id: "meeting:refinement",
        title: "HubSpot Refinement",
        start: `${date}T15:00:00.000Z`,
        end: `${date}T16:00:00.000Z`,
        showAs: "busy"
      }
    ];
  }

  async listUnreadMail(): Promise<MailMessage[]> {
    return [
      {
        id: "mail:1",
        subject: "Test-case walkthrough",
        sender: "Chuck Walter",
        preview: "Would you prefer a recording or a live demo?",
        receivedAt: "2026-06-29T03:31:00.000Z"
      }
    ];
  }

  async syncPlannerBlocks(blocks: WorkBlock[]): Promise<CalendarSyncResponse> {
    this.synced.push(...blocks);
    return { created: blocks.length, updated: 0, skippedMeetings: 2 };
  }
}

export class DemoTeamsAdapter implements TeamsPort {
  async listRecentMessages(): Promise<TeamsMessage[]> {
    return [
      {
        id: "teams:1",
        author: "Joan Chin",
        content: "Please clarify the DTC-5238 acceptance criteria before refinement.",
        createdAt: "2026-06-29T12:00:00.000Z"
      }
    ];
  }
}
