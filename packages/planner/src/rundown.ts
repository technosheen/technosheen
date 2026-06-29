import type {
  DailyRundown,
  JiraIssue,
  MailMessage,
  Meeting,
  ScheduleItem,
  TeamsMessage
} from "@workday/contracts";
import { rankIssues } from "./priority.js";
import { scheduleWork } from "./scheduler.js";
import { generateTimesheet } from "./timesheet.js";

export interface DailyInputs {
  date: string;
  issues: JiraIssue[];
  meetings: Meeting[];
  mail: MailMessage[];
  teams: TeamsMessage[];
  dayStart: string;
  dayEnd: string;
}

export function createDailyRundown(input: DailyInputs): DailyRundown {
  const priorities = rankIssues(input.issues, input.date);
  const workBlocks = scheduleWork(priorities, input.meetings, {
    dayStart: input.dayStart,
    dayEnd: input.dayEnd
  });
  const schedule: ScheduleItem[] = [
    ...input.meetings.map((meeting) => ({ ...meeting, kind: "meeting" as const })),
    ...workBlocks.map((block) => ({ ...block, kind: "work" as const }))
  ].sort((a, b) => a.start.localeCompare(b.start));

  const blockers = priorities
    .filter(({ issue }) => issue.blocked)
    .map(({ issue }) => `${issue.key}: ${issue.summary}`);
  const actionItems = [
    ...priorities
      .filter(({ issue }) => issue.mentionsCurrentUser)
      .map(({ issue }) => `Respond on ${issue.key}: ${issue.summary}`),
    ...input.mail.slice(0, 3).map((message) => `Review email from ${message.sender}: ${message.subject}`),
    ...input.teams.slice(0, 2).map((message) => `Follow up with ${message.author}: ${message.content.slice(0, 90)}`)
  ];

  return {
    date: input.date,
    summary: `${priorities.length} active priorities, ${blockers.length} blockers, ${input.meetings.length} meetings, and ${workBlocks.length} protected focus blocks.`,
    priorities,
    blockers,
    actionItems,
    schedule,
    timesheet: generateTimesheet(schedule)
  };
}
