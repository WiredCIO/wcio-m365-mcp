/**
 * Outlook mail tools.
 *
 * Security note: send_email is one of the highest-risk tools in the package
 * because emails leave the org and can't be unsent. The orchestrating Claude
 * client is expected to confirm with the user before invoking it. We do not
 * implement any "approval workflow" at the MCP layer — that's the client's job.
 */
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graph, GraphError } from "../graph";

const emailAddressSchema = z.object({
  address: z.string().email(),
  name: z.string().optional(),
});

const recipientsField = z
  .array(z.union([z.string().email(), emailAddressSchema]))
  .default([])
  .describe("Recipient list. Accepts either bare email strings or {address, name} objects.");

function toRecipients(recipients: Array<string | { address: string; name?: string }>) {
  return recipients.map((r) =>
    typeof r === "string"
      ? { emailAddress: { address: r } }
      : { emailAddress: { address: r.address, name: r.name } }
  );
}

export function registerMailTools(server: McpServer) {
  // ==========================================================================
  // send_email
  // ==========================================================================
  server.tool(
    "send_email",
    "Send an email immediately from the authenticated user's Outlook mailbox. " +
      "Once sent, the email cannot be recalled programmatically — the client " +
      "calling this tool should confirm with the user before invoking.",
    {
      to: recipientsField,
      cc: recipientsField,
      bcc: recipientsField,
      subject: z.string().describe("Subject line"),
      body_html: z.string().optional().describe("HTML body. Mutually exclusive with body_text."),
      body_text: z.string().optional().describe("Plain-text body. Mutually exclusive with body_html."),
      reply_to: z
        .array(z.string().email())
        .optional()
        .describe("Reply-To addresses (rarely needed)"),
      importance: z.enum(["low", "normal", "high"]).default("normal"),
      save_to_sent_items: z
        .boolean()
        .default(true)
        .describe("Whether to save a copy in the Sent Items folder"),
    },
    async (
      { to, cc, bcc, subject, body_html, body_text, reply_to, importance, save_to_sent_items },
      { authInfo }
    ) => {
      if (!body_html && !body_text) {
        return {
          content: [{ type: "text", text: "Error: must provide either body_html or body_text" }],
          isError: true,
        };
      }
      if (body_html && body_text) {
        return {
          content: [
            { type: "text", text: "Error: provide only one of body_html or body_text, not both" },
          ],
          isError: true,
        };
      }
      if (to.length === 0) {
        return {
          content: [{ type: "text", text: "Error: at least one recipient is required" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;

      const message = {
        subject,
        body: {
          contentType: body_html ? "html" : "text",
          content: body_html ?? body_text,
        },
        toRecipients: toRecipients(to),
        ccRecipients: cc.length ? toRecipients(cc) : undefined,
        bccRecipients: bcc.length ? toRecipients(bcc) : undefined,
        replyTo: reply_to?.length
          ? reply_to.map((addr) => ({ emailAddress: { address: addr } }))
          : undefined,
        importance,
      };

      try {
        await graph(token, "/me/sendMail", {
          method: "POST",
          body: JSON.stringify({
            message,
            saveToSentItems: save_to_sent_items,
          }),
        });
        const recipientSummary = to
          .map((r) => (typeof r === "string" ? r : r.address))
          .join(", ");
        return {
          content: [
            {
              type: "text",
              text: `Sent to ${recipientSummary}. Subject: ${subject}`,
            },
          ],
        };
      } catch (err) {
        return errorResult("send_email", err);
      }
    }
  );

  // ==========================================================================
  // create_draft
  // ==========================================================================
  server.tool(
    "create_draft",
    "Create a draft email in the user's Drafts folder without sending. The user " +
      "can review and send it from Outlook. Lower risk than send_email since " +
      "nothing leaves the mailbox until the user clicks send.",
    {
      to: recipientsField,
      cc: recipientsField,
      bcc: recipientsField,
      subject: z.string(),
      body_html: z.string().optional(),
      body_text: z.string().optional(),
      importance: z.enum(["low", "normal", "high"]).default("normal"),
    },
    async ({ to, cc, bcc, subject, body_html, body_text, importance }, { authInfo }) => {
      if (!body_html && !body_text) {
        return {
          content: [{ type: "text", text: "Error: provide body_html or body_text" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;

      const draft = {
        subject,
        body: {
          contentType: body_html ? "html" : "text",
          content: body_html ?? body_text,
        },
        toRecipients: toRecipients(to),
        ccRecipients: cc.length ? toRecipients(cc) : undefined,
        bccRecipients: bcc.length ? toRecipients(bcc) : undefined,
        importance,
      };

      try {
        const result = (await graph(token, "/me/messages", {
          method: "POST",
          body: JSON.stringify(draft),
        })) as { id: string; webLink?: string };
        return {
          content: [
            {
              type: "text",
              text: `Draft created (id: ${result.id}).${
                result.webLink ? `\nOpen: ${result.webLink}` : ""
              }`,
            },
          ],
        };
      } catch (err) {
        return errorResult("create_draft", err);
      }
    }
  );

  // ==========================================================================
  // reply_to_email
  // ==========================================================================
  server.tool(
    "reply_to_email",
    "Reply to an email message. The reply preserves the thread and quotes the " +
      "original message body automatically.",
    {
      message_id: z
        .string()
        .describe("Message ID of the email being replied to (from outlook_email_search)"),
      comment_html: z
        .string()
        .optional()
        .describe("HTML body of the reply. Either this or comment_text required."),
      comment_text: z.string().optional(),
      reply_all: z.boolean().default(false).describe("Reply to all recipients vs just sender"),
    },
    async ({ message_id, comment_html, comment_text, reply_all }, { authInfo }) => {
      if (!comment_html && !comment_text) {
        return {
          content: [{ type: "text", text: "Error: provide comment_html or comment_text" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;
      const endpoint = reply_all
        ? `/me/messages/${message_id}/replyAll`
        : `/me/messages/${message_id}/reply`;

      try {
        await graph(token, endpoint, {
          method: "POST",
          body: JSON.stringify({
            message: {
              body: {
                contentType: comment_html ? "html" : "text",
                content: comment_html ?? comment_text,
              },
            },
          }),
        });
        return {
          content: [
            { type: "text", text: `Reply${reply_all ? "-all" : ""} sent for message ${message_id}.` },
          ],
        };
      } catch (err) {
        return errorResult("reply_to_email", err);
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
