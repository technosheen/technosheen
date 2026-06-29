import type {
  CalendarSyncResponse,
  JiraIssue,
  MailMessage,
  Meeting,
  TeamsMessage,
  WorkBlock
} from "@workday/contracts";

export interface JiraPort {
  listActiveAssignedIssues(): Promise<JiraIssue[]>;
  listRecentlyUpdatedIssues(since: string): Promise<JiraIssue[]>;
  getIssueDetails(key: string): Promise<{
    issue: JiraIssue;
    comments: Array<{ author: string; body: string; createdAt: string }>;
  }>;
}

export interface OutlookPort {
  listMeetings(start: string, end: string): Promise<Meeting[]>;
  listUnreadMail(since: string): Promise<MailMessage[]>;
  syncPlannerBlocks(blocks: WorkBlock[]): Promise<CalendarSyncResponse>;
}

export interface TeamsPort {
  listRecentMessages(since: string): Promise<TeamsMessage[]>;
}
