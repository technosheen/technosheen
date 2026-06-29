# Workday Planner MVP

AI-assisted workday planning across Jira and explicitly captured work context. The monorepo produces a concise daily rundown, deterministic conflict-free focus blocks, and an exactly eight-hour timesheet that the user can enter into their system of record.

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
- No Microsoft administrator approval for the default explicit-capture mode

## Setup

```bash
cp .env.example .env
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000). The API runs at [http://localhost:4000](http://localhost:4000).

`PLANNER_USE_DEMO_DATA=true` is the safe default for Jira. Set it to `false` and provide Jira credentials for live priorities. Microsoft access is not required: users add only the meetings, emails, or conversations they want considered, then export the generated timesheet as CSV.

## Explicit capture (default)

`MICROSOFT_AUTH_MODE=manual` performs no mailbox scan, Teams scan, or calendar write. Captured context is sent with one rundown request and is not persisted by the MVP.

The dashboard supports:

- Explicit meeting, email, and conversation capture
- On-demand plan generation
- Exact eight-hour timesheet CSV export
- No automatic writes to Outlook or timesheet software

The typed Graph adapters remain optional for organizations that approve them. Set `MICROSOFT_AUTH_MODE=delegated` or `client_credentials` and configure the documented Microsoft variables in `.env`.

## API

### `POST /api/rundown/daily`

Optional body, including only context the user chose to capture:

```json
{
  "date": "2026-06-29",
  "captures": {
    "meetings": [],
    "mail": [],
    "teams": []
  }
}
```

Returns `summary`, scored `priorities`, `blockers`, `actionItems`, a meeting-safe `schedule`, and an exactly 480-minute `timesheet`.

### `POST /api/automation/daily`

Uses the same request and response contract as `/api/rundown/daily`. Configure `PLANNER_AUTOMATION_TOKEN` and send it as `Authorization: Bearer <token>`. This enables Task Scheduler, cron, Shortcuts, or another user-controlled trigger without granting mailbox access.

Example:

```bash
curl -X POST http://localhost:4000/api/automation/daily \
  -H "Authorization: Bearer $PLANNER_AUTOMATION_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{}'
```

`POST /api/calendar/sync` and the Microsoft connection endpoints remain available only for optional managed Microsoft integrations; the default dashboard never calls them.

## Scheduling rules

- `HUB-*` calendar title: `HUBSPOT | HUB-*`
- `DTC-*` calendar title: `DTC | DTC-*`
- Meetings and internal timesheet activity: `INT-58`
- Captured meetings are immutable busy intervals.
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
