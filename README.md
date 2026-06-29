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
- Microsoft Entra app registration with delegated permissions (no client secret required)

## Setup

```bash
cp .env.example .env
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000). The API runs at [http://localhost:4000](http://localhost:4000).

`PLANNER_USE_DEMO_DATA=true` is the safe default. It exercises the full planner and calendar-sync flow without external writes. To use live integrations, set it to `false` and provide all Jira and Microsoft Graph variables in `.env`.

## Microsoft 365 delegated sign-in

Delegated device sign-in is the default. It opens Microsoft's own login page, operates only as the signed-in user, and stores refresh tokens in the operating system's encrypted credential store. The planner never receives a password or MFA code.

1. Create an Entra app registration for organizational accounts. You still need a client/application ID, but not a client secret.
2. Under **Authentication → Advanced settings**, enable **Allow public client flows**.
3. Add delegated Microsoft Graph permissions:
   - `User.Read`
   - `Calendars.ReadWrite`
   - `Mail.Read`
   - `Chat.Read`
4. Set `MICROSOFT_CLIENT_ID` and leave `MICROSOFT_TENANT_ID=organizations`, or use your tenant ID/domain to restrict sign-in.
5. Set `PLANNER_USE_DEMO_DATA=false`, start the app, and choose **Connect Microsoft**.

These delegated permissions normally support user consent, although a company can disable user consent or require approval. Channel-wide Teams messages are excluded by default because delegated `ChannelMessage.Read.All` requires administrator consent. Set `MICROSOFT_ENABLE_TEAMS_CHANNELS=true` only after that permission is approved.

The legacy app-only path remains available with `MICROSOFT_AUTH_MODE=client_credentials`, `MICROSOFT_CLIENT_SECRET`, and `MICROSOFT_USER_ID`. It requires administrator-approved application permissions.

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

### Microsoft connection endpoints

- `GET /api/auth/microsoft/status`
- `POST /api/auth/microsoft/device/start`
- `GET /api/auth/microsoft/device/:sessionId`
- `POST /api/auth/microsoft/disconnect`

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
