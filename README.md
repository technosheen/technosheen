# Workday Planner MVP

AI-assisted workday planning across Jira, Outlook Calendar/Mail, and Microsoft Teams. The monorepo produces a concise daily rundown, deterministic conflict-free focus blocks, planner-owned Outlook events, and an exactly eight-hour timesheet.

## Architecture

```text
apps/
  api/              Fastify HTTP API and application composition
  web/              Next.js dashboard
packages/
  contracts/        Shared Zod request/response and domain schemas
  planner/          Pure priority, scheduling, overlap, and timesheet logic
  integrations/     Typed Jira and Microsoft Graph adapters
prisma/
  schema.prisma     Snapshot, plan, sync-run, and timesheet persistence model
```

Dependencies point inward: adapters implement ports, the API composes use cases, and the planner package contains no network or framework code.

## Requirements

- Node.js 22+
- npm 11+
- Jira Cloud API token for live Jira access
- Microsoft Entra application with application permissions for Calendar, Mail, and Teams reads/writes

## Setup

```bash
cp .env.example .env
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000). The API runs at [http://localhost:4000](http://localhost:4000).

`PLANNER_USE_DEMO_DATA=true` is the safe default. It exercises the full planner and calendar-sync flow without external writes. To use live integrations, set it to `false` and provide all Jira and Microsoft Graph variables in `.env`.

## Microsoft Graph permissions

The client-credentials adapter expects application permissions approved by an Entra administrator:

- `Calendars.ReadWrite`
- `Mail.Read`
- `Chat.Read.All`
- `ChannelMessage.Read.All`
- `Team.ReadBasic.All`
- `User.Read.All`

Set `MICROSOFT_USER_ID` to the target user's UPN or Graph user ID. Secrets remain environment-only and must never be committed.

## API

### `POST /api/rundown/daily`

Optional body:

```json
{ "date": "2026-06-29" }
```

Returns `summary`, scored `priorities`, `blockers`, `actionItems`, a meeting-safe `schedule`, and an exactly 480-minute `timesheet`.

### `POST /api/calendar/sync`

```json
{ "schedule": [] }
```

Only `kind: "work"` items are synchronized. Planner events carry an HTML ownership marker and are created with `showAs: "free"`. Existing events are updated only when that marker matches the stable planner block ID. Human-owned Outlook meetings remain busy and are never updated, duplicated, or deleted.

## Scheduling rules

- `HUB-*` calendar title: `HUBSPOT | HUB-*`
- `DTC-*` calendar title: `DTC | DTC-*`
- Meetings and internal timesheet activity: `INT-58`
- Outlook meetings are immutable busy intervals.
- Planner work blocks are free and receive a 15-minute meeting buffer.
- Priority scoring combines Jira priority, workflow status, due date, mentions, and blockers.
- Timesheet generation includes meetings and Jira work, then balances remaining time to `INT-58` so the total is exactly eight hours.

## Validation

```bash
npm test
npm run typecheck
npm run build
```

Tests cover title prefixes, overlap detection, meeting buffers, Outlook ownership/free-busy behavior, exact timesheet balancing, and both API acceptance routes.
