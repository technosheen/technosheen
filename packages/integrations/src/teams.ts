import type { TeamsMessage } from "@workday/contracts";
import type { TeamsPort } from "./ports.js";
import type { GraphClient } from "./graph.js";

interface GraphCollection<T> {
  value: T[];
}

interface GraphChat {
  id: string;
}

interface GraphChatMessage {
  id: string;
  createdDateTime: string;
  body: { content: string };
  from?: { user?: { displayName?: string } };
  webUrl?: string;
}

interface GraphTeam {
  id: string;
}

interface GraphChannel {
  id: string;
}

export class TeamsAdapter implements TeamsPort {
  constructor(
    private readonly graph: Pick<GraphClient, "request">,
    private readonly userId: string
  ) {}

  async listRecentMessages(since: string): Promise<TeamsMessage[]> {
    const [chatMessages, channelMessages] = await Promise.all([
      this.#listChatMessages(),
      this.#listChannelMessages()
    ]);

    return [...chatMessages, ...channelMessages]
      .filter((message) => message.createdDateTime >= since)
      .map((message) => ({
        id: message.id,
        author: message.from?.user?.displayName ?? "Unknown",
        content: stripHtml(message.body.content),
        createdAt: new Date(message.createdDateTime).toISOString(),
        ...(message.webUrl ? { webUrl: message.webUrl } : {})
      }))
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  async #listChatMessages(): Promise<GraphChatMessage[]> {
    const chats = await this.graph.request<GraphCollection<GraphChat>>(
      `/users/${encodeURIComponent(this.userId)}/chats?$top=25`
    );
    const results = await Promise.all(
      chats.value.map(async (chat) => {
        const messages = await this.graph.request<GraphCollection<GraphChatMessage>>(
          `/chats/${encodeURIComponent(chat.id)}/messages?$top=25`
        );
        return messages.value;
      })
    );
    return results.flat();
  }

  async #listChannelMessages(): Promise<GraphChatMessage[]> {
    const teams = await this.graph.request<GraphCollection<GraphTeam>>(
      `/users/${encodeURIComponent(this.userId)}/joinedTeams`
    );
    const byTeam = await Promise.all(
      teams.value.map(async (team) => {
        const channels = await this.graph.request<GraphCollection<GraphChannel>>(
          `/teams/${encodeURIComponent(team.id)}/channels`
        );
        const byChannel = await Promise.all(
          channels.value.map(async (channel) => {
            const messages = await this.graph.request<GraphCollection<GraphChatMessage>>(
              `/teams/${encodeURIComponent(team.id)}/channels/${encodeURIComponent(channel.id)}/messages?$top=25`
            );
            return messages.value;
          })
        );
        return byChannel.flat();
      })
    );
    return byTeam.flat();
  }
}

function stripHtml(value: string): string {
  return value.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}
