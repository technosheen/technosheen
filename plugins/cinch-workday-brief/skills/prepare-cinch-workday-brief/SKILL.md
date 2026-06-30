---
name: prepare-cinch-workday-brief
description: Generate a deterministic Cinch workday plan and exactly eight-hour timesheet from the authenticated user's approved Outlook, Jira, Microsoft Teams, and GitHub connectors. Use for a morning brief, daily plan, schedule, priorities, blockers, or timesheet.
---

# Prepare Cinch Workday Plan

Use the user's existing Codex connectors for retrieval and the bundled
`generate_workday_plan` tool for deterministic scheduling and timesheet
balancing. The planner itself must not request Microsoft Graph credentials.

## Authorization boundary

- Retrieve only data the authenticated user can access through connected apps.
- Never ask for passwords, access tokens, session cookies, or Microsoft
  application credentials.
- Do not use the planner's optional Graph adapters.
- Do not send messages, modify Jira, write calendar events, or submit a
  timesheet.
- Select only context needed for today's plan. Do not perform an unrestricted
  mailbox or Teams export.

If a connector is unavailable, continue with authorized sources and name the
missing source under `Source limitations`.

## Workflow

1. Resolve today's date and current time in `America/New_York`.
2. Read today's calendar through the Outlook Calendar connector. Capture real
   meetings as immutable busy intervals with stable IDs, ISO timestamps, and
   links when available.
3. Select high-signal unread Outlook mail: explicit requests, deadlines,
   blockers, or messages likely requiring a reply.
4. Search Jira for assigned, mentioned, due, overdue, blocked, in-progress,
   in-review, recently updated, high-priority incident, and bug work related to
   Cinch. The planner API owns Jira priority scoring when its Jira adapter is
   configured; use connector results to explain context and limitations.
5. Select only Teams messages that materially affect a priority, blocker,
   commitment, or meeting. Use GitHub only for relevant review or CI context.
6. Deduplicate repeated work across sources.
7. Call `generate_workday_plan` once with the selected meetings, mail, Teams
   messages, and date. Preserve its schedule and timesheet totals exactly.
8. If the tool reports that the local API is unavailable, tell the user to run
   `npm run dev` in the Workday Planner checkout or configure
   `WORKDAY_PLANNER_API_URL`. Do not silently replace deterministic output with
   an estimated schedule.

## Output

Use these sections in order:

### Calendar

List real meetings and generated focus blocks chronologically. Identify
conflicts, preparation needs, and notable gaps. Real meetings are busy;
generated work blocks are free.

### Important Unread Email

Give sender, subject, reason, likely action, and link when available.

### Jira / Cinch Tickets

Give ticket key, owner, status, priority or due date, reason for attention, and
link when available.

### Attention Today

Give a short ordered checklist, then the generated timesheet entries and exact
eight-hour total.

Keep the result concise, distinguish facts from inference, and include a
one-line `Source limitations` note when any expected source was unavailable.
