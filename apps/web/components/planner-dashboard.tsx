"use client";

import { useEffect, useState, useTransition } from "react";
import type {
  DailyRundown,
  MicrosoftAuthStatus,
  MicrosoftDeviceLogin,
  MicrosoftDeviceLoginStatus,
  ScheduleItem
} from "@workday/contracts";

const API_URL = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:4000";
const APP_TIME_ZONE = process.env.NEXT_PUBLIC_PLANNER_TIME_ZONE ?? "America/New_York";

export function PlannerDashboard({ dateLabel }: { dateLabel: string }) {
  const [rundown, setRundown] = useState<DailyRundown | null>(null);
  const [status, setStatus] = useState("Ready to build your day");
  const [microsoft, setMicrosoft] = useState<MicrosoftAuthStatus | null>(null);
  const [deviceLogin, setDeviceLogin] = useState<MicrosoftDeviceLogin | null>(null);
  const [isPending, startTransition] = useTransition();

  useEffect(() => {
    void refreshMicrosoftStatus();
  }, []);

  async function refreshMicrosoftStatus() {
    const response = await fetch(`${API_URL}/api/auth/microsoft/status`);
    if (response.ok) setMicrosoft((await response.json()) as MicrosoftAuthStatus);
  }

  async function connectMicrosoft() {
    setStatus("Starting secure Microsoft sign-in…");
    const response = await fetch(`${API_URL}/api/auth/microsoft/device/start`, { method: "POST" });
    if (!response.ok) {
      setStatus("Microsoft sign-in could not be started");
      return;
    }
    const login = (await response.json()) as MicrosoftDeviceLogin;
    setDeviceLogin(login);
    setStatus("Complete sign-in with Microsoft");
    void pollDeviceLogin(login.sessionId);
  }

  async function pollDeviceLogin(sessionId: string) {
    for (let attempt = 0; attempt < 90; attempt += 1) {
      await new Promise((resolve) => window.setTimeout(resolve, 2_000));
      const response = await fetch(`${API_URL}/api/auth/microsoft/device/${sessionId}`);
      if (!response.ok) continue;
      const result = (await response.json()) as MicrosoftDeviceLoginStatus;
      if (result.status === "pending") continue;
      if (result.status === "connected") {
        setDeviceLogin(null);
        setStatus(`Microsoft connected as ${result.account?.username ?? "your account"}`);
        await refreshMicrosoftStatus();
        return;
      }
      setStatus(result.error ?? "Microsoft sign-in failed");
      return;
    }
    setStatus("Microsoft sign-in expired; start it again");
  }

  async function loadRundown() {
    setStatus("Reviewing Jira, Outlook, and Teams…");
    const response = await fetch(`${API_URL}/api/rundown/daily`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({})
    });
    if (!response.ok) {
      setStatus("Could not generate the rundown");
      return;
    }
    const data = (await response.json()) as DailyRundown;
    startTransition(() => {
      setRundown(data);
      setStatus("Daily rundown ready");
    });
  }

  async function syncCalendar() {
    if (!rundown) return;
    setStatus("Syncing planner-owned blocks…");
    const response = await fetch(`${API_URL}/api/calendar/sync`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ schedule: rundown.schedule })
    });
    const result = (await response.json()) as { created: number; updated: number };
    setStatus(`Outlook synced: ${result.created} created, ${result.updated} updated`);
  }

  return (
    <main className="app-shell">
      <Sidebar />
      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="date-label">{dateLabel}</p>
            <h1>Plan the workday</h1>
          </div>
          <div className="topbar-actions">
            {microsoft?.mode === "delegated" && microsoft.status === "disconnected" ? (
              <button className="button secondary" onClick={connectMicrosoft}>
                Connect Microsoft
              </button>
            ) : null}
            {microsoft?.account ? (
              <span className="connection-badge">Microsoft · {microsoft.account.username}</span>
            ) : null}
            <span className="sync-status" aria-live="polite">{status}</span>
            <button
              className="button secondary"
              onClick={loadRundown}
              disabled={isPending || microsoft?.status === "disconnected"}
            >
              Generate rundown
            </button>
            <button className="button primary" onClick={syncCalendar} disabled={!rundown}>
              Sync to Outlook
            </button>
          </div>
        </header>

        {deviceLogin ? <MicrosoftSignIn login={deviceLogin} /> : null}
        <Summary rundown={rundown} />

        <section className="planner-grid">
          <PriorityQueue rundown={rundown} />
          <Schedule rundown={rundown} />
          <ActionItems rundown={rundown} />
        </section>

        <Timesheet rundown={rundown} />
      </section>
    </main>
  );
}

function MicrosoftSignIn({ login }: { login: MicrosoftDeviceLogin }) {
  return (
    <section className="microsoft-signin" aria-live="polite">
      <div>
        <strong>Connect your work Microsoft account</strong>
        <span>Microsoft will ask for this one-time code. Credentials never pass through the planner.</span>
      </div>
      <code>{login.userCode}</code>
      <a
        className="button primary"
        href={login.verificationUri}
        target="_blank"
        rel="noreferrer"
      >
        Open Microsoft sign-in
      </a>
    </section>
  );
}

function Sidebar() {
  return (
    <aside className="sidebar">
      <div className="wordmark">
        <span className="logo-mark">✓</span>
        <span>Workday<br /><strong>Planner</strong></span>
      </div>
      <nav aria-label="Primary">
        {["Today", "Schedule", "Timesheet", "Integrations"].map((item, index) => (
          <a key={item} href={`#${item.toLowerCase()}`} className={index === 0 ? "active" : ""}>
            <span aria-hidden="true">{["▣", "□", "◷", "◇"][index]}</span>{item}
          </a>
        ))}
      </nav>
      <div className="sidebar-footer">
        <span className="avatar">SM</span>
        <span>Sean Mahoney<small>Engineering</small></span>
      </div>
    </aside>
  );
}

function Summary({ rundown }: { rundown: DailyRundown | null }) {
  const meetings = rundown?.schedule.filter((item) => item.kind === "meeting").length ?? 0;
  const focusMinutes =
    rundown?.schedule
      .filter((item) => item.kind === "work")
      .reduce((total, item) => total + minutesBetween(item), 0) ?? 0;
  const items = [
    ["AI Daily Rundown", rundown?.summary ?? "Generate a rundown to review today’s work."],
    ["Focus Time Available", `${formatDuration(focusMinutes)} across protected blocks`],
    ["Meetings", `${meetings} busy events protected`],
    ["Timesheet Preview", rundown ? "8h 00m planned · balanced" : "Waiting for schedule"]
  ];
  return (
    <section className="summary-band" aria-label="Daily summary">
      {items.map(([label, value]) => (
        <div key={label}>
          <strong>{label}</strong>
          <span>{value}</span>
        </div>
      ))}
    </section>
  );
}

function PriorityQueue({ rundown }: { rundown: DailyRundown | null }) {
  return (
    <section className="panel priority-panel" id="today">
      <PanelHeader title="Priority Queue" meta="AI prioritized" />
      <div className="priority-head">
        <span>Key</span><span>Summary</span><span>Status</span><span>Score</span>
      </div>
      {(rundown?.priorities ?? []).map(({ issue, score }) => (
        <div className="priority-row" key={issue.key}>
          <a href={`https://cinchhs.atlassian.net/browse/${issue.key}`} target="_blank" rel="noreferrer">
            {issue.key}
          </a>
          <span>{issue.summary}</span>
          <span>{issue.status}</span>
          <strong className={score >= 80 ? "score high" : "score"}>{score}</strong>
        </div>
      ))}
      {!rundown ? <EmptyState label="No priorities loaded" /> : null}
    </section>
  );
}

function Schedule({ rundown }: { rundown: DailyRundown | null }) {
  return (
    <section className="panel schedule-panel" id="schedule">
      <PanelHeader title="Day Schedule" meta="9 AM – 5 PM" />
      <div className="legend"><span className="busy-key">Busy · Outlook</span><span className="free-key">Free · Planner</span></div>
      <div className="timeline">
        <div className="time-axis">
          {["9 AM", "10 AM", "11 AM", "12 PM", "1 PM", "2 PM", "3 PM", "4 PM", "5 PM"].map((time) => <span key={time}>{time}</span>)}
        </div>
        <div className="timeline-events">
          {(rundown?.schedule ?? []).map((item) => <ScheduleBlock key={item.id} item={item} />)}
          {!rundown ? <EmptyState label="Generate the rundown to place work around meetings" /> : null}
        </div>
      </div>
    </section>
  );
}

function ScheduleBlock({ item }: { item: ScheduleItem }) {
  const startHour = hourInTimeZone(item.start);
  const endHour = hourInTimeZone(item.end);
  const top = ((startHour - 9) / 8) * 100;
  const height = Math.max(((endHour - startHour) / 8) * 100, 6);
  return (
    <article className={`schedule-block ${item.kind}`} style={{ top: `${top}%`, height: `${height}%` }}>
      <strong>{item.title}</strong>
      <span>{formatTime(item.start)} – {formatTime(item.end)}</span>
      <small>{item.kind === "meeting" ? "Busy" : "Planner · Free"}</small>
    </article>
  );
}

function ActionItems({ rundown }: { rundown: DailyRundown | null }) {
  return (
    <section className="panel actions-panel">
      <PanelHeader title="Action Items" meta={String(rundown?.actionItems.length ?? 0)} />
      <ul>
        {(rundown?.actionItems ?? []).map((item) => (
          <li key={item}><input type="checkbox" aria-label={`Complete ${item}`} /><span>{item}</span></li>
        ))}
      </ul>
      {!rundown ? <EmptyState label="No action items loaded" /> : null}
    </section>
  );
}

function Timesheet({ rundown }: { rundown: DailyRundown | null }) {
  return (
    <section className="timesheet panel" id="timesheet">
      <PanelHeader title="Timesheet Preview" meta="Must total 8h 00m" />
      <div className="timesheet-head">
        <span>Code</span><span>Activity</span><span>Source</span><span>Duration</span>
      </div>
      {(rundown?.timesheet.entries ?? []).map((entry) => (
        <div className="timesheet-row" key={entry.id}>
          <strong>{entry.code}</strong><span>{entry.description}</span><span>{entry.source}</span><span>{formatDuration(entry.minutes)}</span>
        </div>
      ))}
      <footer><span>✓ Total is exactly 8h 00m</span><strong>Total: 8h 00m</strong></footer>
    </section>
  );
}

function PanelHeader({ title, meta }: { title: string; meta: string }) {
  return <header className="panel-header"><h2>{title}</h2><span>{meta}</span></header>;
}

function EmptyState({ label }: { label: string }) {
  return <div className="empty-state">{label}</div>;
}

function minutesBetween(item: ScheduleItem) {
  return Math.round((new Date(item.end).getTime() - new Date(item.start).getTime()) / 60_000);
}

function formatDuration(minutes: number) {
  return `${Math.floor(minutes / 60)}h ${String(minutes % 60).padStart(2, "0")}m`;
}

function formatTime(value: string) {
  return new Intl.DateTimeFormat("en-US", {
    hour: "numeric",
    minute: "2-digit",
    timeZone: APP_TIME_ZONE
  }).format(new Date(value));
}

function hourInTimeZone(value: string) {
  const parts = new Intl.DateTimeFormat("en-US", {
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
    timeZone: APP_TIME_ZONE
  }).formatToParts(new Date(value));
  const values = Object.fromEntries(
    parts.filter((part) => part.type !== "literal").map((part) => [part.type, Number(part.value)])
  );
  return (values.hour ?? 0) + (values.minute ?? 0) / 60;
}
