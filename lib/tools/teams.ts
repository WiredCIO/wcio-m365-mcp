/**
 * Microsoft Teams messaging tools.
 *
 * Teams has two distinct concepts:
 *  - Chats: 1:1 and group chats, identified by chatId
 *  - Channels: messages in a Teams team channel, identified by teamId + channelId
 *
 * Use the read-only Microsoft 365 connector's chat_message_search to discover
 * IDs before calling these tools.
 */
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graph, GraphError } from "../graph";

export function registerTeamsTools(server: McpServer) {
  // ==========================================================================
  // post_chat_message — 1:1 or group chat
  // ==========================================================================
  server.tool(
    "post_chat_message",
    "Post a message into a Teams 1:1 or group chat. Use chat_message_search " +
      "on the read-only connector to find chat IDs first. For posting to a " +
      "channel of a team, use post_channel_message instead.",
    {
      chat_id: z.string().describe("Teams chat ID, e.g. 19:xxx@thread.v2"),
      content_html: z
        .string()
        .optional()
        .describe("HTML message body. Supports basic formatting, links, @mentions."),
      content_text: z.string().optional().describe("Plain text message body"),
      importance: z.enum(["normal", "high", "urgent"]).default("normal"),
    },
    async ({ chat_id, content_html, content_text, importance }, { authInfo }) => {
      if (!content_html && !content_text) {
        return {
          content: [{ type: "text", text: "Error: provide content_html or content_text" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;

      try {
        const result = (await graph(token, `/chats/${chat_id}/messages`, {
          method: "POST",
          body: JSON.stringify({
            body: {
              contentType: content_html ? "html" : "text",
              content: content_html ?? content_text,
            },
            importance,
          }),
        })) as { id: string; webUrl?: string };
        return {
          content: [
            {
              type: "text",
              text: `Posted to chat ${chat_id} (message id ${result.id}).${
                result.webUrl ? `\nLink: ${result.webUrl}` : ""
              }`,
            },
          ],
        };
      } catch (err) {
        return errorResult("post_chat_message", err);
      }
    }
  );

  // ==========================================================================
  // post_channel_message — team channel
  // ==========================================================================
  server.tool(
    "post_channel_message",
    "Post a message into a Teams channel. Requires the team ID and channel ID.",
    {
      team_id: z.string().describe("Teams team ID"),
      channel_id: z.string().describe("Channel ID within that team"),
      content_html: z.string().optional(),
      content_text: z.string().optional(),
      subject: z.string().optional().describe("Optional subject heading for the post"),
      importance: z.enum(["normal", "high", "urgent"]).default("normal"),
    },
    async (
      { team_id, channel_id, content_html, content_text, subject, importance },
      { authInfo }
    ) => {
      if (!content_html && !content_text) {
        return {
          content: [{ type: "text", text: "Error: provide content_html or content_text" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;

      const message: Record<string, unknown> = {
        body: {
          contentType: content_html ? "html" : "text",
          content: content_html ?? content_text,
        },
        importance,
      };
      if (subject) message.subject = subject;

      try {
        const result = (await graph(
          token,
          `/teams/${team_id}/channels/${channel_id}/messages`,
          {
            method: "POST",
            body: JSON.stringify(message),
          }
        )) as { id: string; webUrl?: string };
        return {
          content: [
            {
              type: "text",
              text: `Posted to channel (message id ${result.id}).${
                result.webUrl ? `\nLink: ${result.webUrl}` : ""
              }`,
            },
          ],
        };
      } catch (err) {
        return errorResult("post_channel_message", err);
      }
    }
  );

  // ==========================================================================
  // reply_to_channel_message
  // ==========================================================================
  server.tool(
    "reply_to_channel_message",
    "Reply to an existing channel message, creating a threaded reply.",
    {
      team_id: z.string(),
      channel_id: z.string(),
      parent_message_id: z.string().describe("ID of the channel message being replied to"),
      content_html: z.string().optional(),
      content_text: z.string().optional(),
    },
    async (
      { team_id, channel_id, parent_message_id, content_html, content_text },
      { authInfo }
    ) => {
      if (!content_html && !content_text) {
        return {
          content: [{ type: "text", text: "Error: provide content_html or content_text" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;

      try {
        const result = (await graph(
          token,
          `/teams/${team_id}/channels/${channel_id}/messages/${parent_message_id}/replies`,
          {
            method: "POST",
            body: JSON.stringify({
              body: {
                contentType: content_html ? "html" : "text",
                content: content_html ?? content_text,
              },
            }),
          }
        )) as { id: string };
        return {
          content: [{ type: "text", text: `Reply posted (id ${result.id}).` }],
        };
      } catch (err) {
        return errorResult("reply_to_channel_message", err);
      }
    }
  );
}

function errorResult(toolName: string, err: unknown) {
  if (err instanceof GraphError) {
    return {
      content: [
        {
          type: "text" as const,
          text:
            `Microsoft Graph error in ${toolName}: ${err.message}\n` +
            `HTTP ${err.statusCode}${err.graphCode ? ` (${err.graphCode})` : ""}`,
        },
      ],
      isError: true,
    };
  }
  const msg = err instanceof Error ? err.message : String(err);
  return {
    content: [{ type: "text" as const, text: `Error in ${toolName}: ${msg}` }],
    isError: true,
  };
}
