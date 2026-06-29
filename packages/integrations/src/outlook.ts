import type {
  CalendarSyncResponse,
  MailMessage,
  Meeting,
  WorkBlock
} from "@workday/contracts";
import type { OutlookPort } from "./ports.js";
import type { GraphClient } from "./graph.js";

const MARKER_PREFIX = "workday-planner:";

interface GraphEvent {
  id: string;
  subject: string;
  body?: { content?: string };
  bodyPreview?: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  showAs?: string;
  webLink?: string;
}

interface GraphMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  sender?: { emailAddress?: { name?: string; address?: string } };
  webLink?: string;
}

interface GraphCollection<T> {
  value: T[];
}

export interface OutlookAdapterOptions {
  graph: Pick<GraphClient, "request">;
  userId: string;
}

export class OutlookAdapter implements OutlookPort {
  readonly #graph: Pick<GraphClient, "request">;
  readonly #userPath: string;

  constructor(options: OutlookAdapterOptions) {
    this.#graph = options.graph;
    this.#userPath = `/users/${encodeURIComponent(options.userId)}`;
  }

  async listMeetings(start: string, end: string): Promise<Meeting[]> {
    const query = new URLSearchParams({
      startDateTime: start,
      endDateTime: end,
      $select: "id,subject,start,end,showAs,webLink,body"
    });
    const response = await this.#graph.request<GraphCollection<GraphEvent>>(
      `${this.#userPath}/calendarView?${query}`,
      { headers: { Prefer: 'outlook.timezone="UTC"' } }
    );
    return response.value
      .filter((event) => !plannerIdFromEvent(event))
      .map((event) => ({
        id: event.id,
        title: event.subject,
        start: asIso(event.start),
        end: asIso(event.end),
        showAs: "busy" as const,
        ...(event.webLink ? { webUrl: event.webLink } : {})
      }));
  }

  async listUnreadMail(since: string): Promise<MailMessage[]> {
    const query = new URLSearchParams({
      $filter: `isRead eq false and receivedDateTime ge ${since}`,
      $orderby: "receivedDateTime desc",
      $top: "50",
      $select: "id,subject,bodyPreview,receivedDateTime,sender,webLink"
    });
    const response = await this.#graph.request<GraphCollection<GraphMessage>>(
      `${this.#userPath}/mailFolders/inbox/messages?${query}`
    );
    return response.value.map((message) => ({
      id: message.id,
      subject: message.subject,
      sender:
        message.sender?.emailAddress?.name ??
        message.sender?.emailAddress?.address ??
        "Unknown sender",
      preview: message.bodyPreview,
      receivedAt: new Date(message.receivedDateTime).toISOString(),
      ...(message.webLink ? { webUrl: message.webLink } : {})
    }));
  }

  async syncPlannerBlocks(blocks: WorkBlock[]): Promise<CalendarSyncResponse> {
    if (blocks.length === 0) return { created: 0, updated: 0, skippedMeetings: 0 };
    const start = blocks.reduce((minimum, block) => (block.start < minimum ? block.start : minimum), blocks[0]!.start);
    const end = blocks.reduce((maximum, block) => (block.end > maximum ? block.end : maximum), blocks[0]!.end);
    const query = new URLSearchParams({
      startDateTime: start,
      endDateTime: end,
      $select: "id,subject,start,end,showAs,body"
    });
    const existing = await this.#graph.request<GraphCollection<GraphEvent>>(
      `${this.#userPath}/calendarView?${query}`,
      { headers: { Prefer: 'outlook.timezone="UTC"' } }
    );
    const plannerEvents = new Map(
      existing.value.flatMap((event) => {
        const plannerId = plannerIdFromEvent(event);
        return plannerId ? [[plannerId, event] as const] : [];
      })
    );

    let created = 0;
    let updated = 0;
    for (const block of blocks) {
      const payload = graphEventFromBlock(block);
      const event = plannerEvents.get(block.id);
      if (event) {
        await this.#graph.request(`${this.#userPath}/events/${encodeURIComponent(event.id)}`, {
          method: "PATCH",
          body: JSON.stringify(payload)
        });
        updated += 1;
      } else {
        await this.#graph.request(`${this.#userPath}/events`, {
          method: "POST",
          body: JSON.stringify(payload)
        });
        created += 1;
      }
    }

    return {
      created,
      updated,
      skippedMeetings: existing.value.length - plannerEvents.size
    };
  }
}

function asIso(value: GraphEvent["start"]): string {
  const hasOffset = /Z$|[+-]\d{2}:\d{2}$/.test(value.dateTime);
  return new Date(hasOffset ? value.dateTime : `${value.dateTime}Z`).toISOString();
}

function plannerIdFromEvent(event: GraphEvent): string | null {
  const content = event.body?.content ?? event.bodyPreview ?? "";
  const match = content.match(/<!-- workday-planner:([^>]+) -->/);
  return match?.[1] ?? null;
}

function graphEventFromBlock(block: WorkBlock) {
  return {
    subject: block.title,
    body: {
      contentType: "html",
      content: `<!-- ${MARKER_PREFIX}${block.id} --><p>Created and managed by Workday Planner.</p>`
    },
    start: { dateTime: block.start, timeZone: "UTC" },
    end: { dateTime: block.end, timeZone: "UTC" },
    showAs: "free",
    isReminderOn: false,
    categories: ["Workday Planner"]
  };
}
