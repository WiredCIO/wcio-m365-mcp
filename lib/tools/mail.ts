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
    "Reply to an email message and SEND it immediately. The reply preserves the " +
      "thread and quotes the original message body automatically. Once sent it " +
      "cannot be recalled, so confirm with the user first. To leave a threaded " +
      "reply in Drafts for the user to review instead, use create_reply_draft.",
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

  // ==========================================================================
  // create_reply_draft
  // ==========================================================================
  // Outlook builds the reply itself (createReply / createReplyAll), so the
  // draft carries the real thread: In-Reply-To and References headers, the
  // conversation id, the "RE:" subject, the original recipients, and the
  // quoted history. We then write our text in above the quoted history and
  // optionally add recipients. Nothing is sent.
  server.tool(
    "create_reply_draft",
    "Create a reply to an existing email and leave it UNSENT in the user's " +
      "Drafts folder, threaded on the original conversation exactly as if the " +
      "user had clicked Reply (or Reply All) in Outlook. The quoted history, " +
      "subject and recipients come from Outlook itself. Your text goes above " +
      "the quoted history. Optionally add To or Cc recipients. The user " +
      "reviews, edits, sends or deletes the draft in Outlook.",
    {
      message_id: z
        .string()
        .describe(
          "Graph message ID of the email being replied to. Usually the latest " +
            "message in the thread (from outlook_email_search)."
        ),
      body_html: z
        .string()
        .optional()
        .describe("HTML for the new reply text. Mutually exclusive with body_text."),
      body_text: z
        .string()
        .optional()
        .describe("Plain text for the new reply text. Mutually exclusive with body_html."),
      reply_all: z
        .boolean()
        .default(false)
        .describe("Reply to everyone on the message (true) or just the sender (false)"),
      add_to: recipientsField.describe("Extra To recipients beyond the ones Outlook fills in"),
      add_cc: recipientsField.describe("Extra Cc recipients beyond the ones Outlook fills in"),
      time_zone: z
        .string()
        .default("Central Standard Time")
        .describe(
          "Windows time zone name used for the dates in the quoted header " +
            "(e.g. 'Eastern Standard Time'). Outlook uses UTC if not set."
        ),
    },
    async (
      { message_id, body_html, body_text, reply_all, add_to, add_cc, time_zone },
      { authInfo }
    ) => {
      if (!body_html && !body_text) {
        return {
          content: [{ type: "text", text: "Error: provide body_html or body_text" }],
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
      const token = authInfo!.extra!.accessToken as string;
      const endpoint = reply_all
        ? `/me/messages/${encodeURIComponent(message_id)}/createReplyAll`
        : `/me/messages/${encodeURIComponent(message_id)}/createReply`;

      // Step 1: let Outlook build the threaded draft.
      let draft: DraftMessage;
      try {
        draft = await graph<DraftMessage>(token, endpoint, {
          method: "POST",
          body: JSON.stringify({}),
          headers: {
            Prefer: `outlook.body-content-type="html", outlook.timezone="${time_zone.replace(/"/g, "")}"`,
          },
        });
      } catch (err) {
        return errorResult("create_reply_draft", err);
      }

      // Step 2: write our text above the quoted history and add recipients.
      const newHtml = body_html ?? textToHtml(body_text!);
      const existing = draft.body?.content ?? "";
      const isHtml = (draft.body?.contentType ?? "html").toLowerCase() === "html";
      const content = isHtml
        ? insertAfterBodyTag(existing, `<div>${newHtml}</div>`)
        : `${body_text ?? htmlToRoughText(body_html!)}\n\n${existing}`;

      const patch: Record<string, unknown> = {
        body: { contentType: isHtml ? "html" : "text", content },
      };
      if (add_to.length) {
        patch.toRecipients = mergeRecipients(draft.toRecipients ?? [], toRecipients(add_to));
      }
      if (add_cc.length) {
        patch.ccRecipients = mergeRecipients(draft.ccRecipients ?? [], toRecipients(add_cc));
      }

      let updated: DraftMessage;
      try {
        updated = await graph<DraftMessage>(token, `/me/messages/${encodeURIComponent(draft.id)}`, {
          method: "PATCH",
          body: JSON.stringify(patch),
        });
      } catch (err) {
        // Don't leave an empty reply sitting in Drafts if we couldn't fill it.
        try {
          await graph(token, `/me/messages/${encodeURIComponent(draft.id)}`, { method: "DELETE" });
        } catch {
          // best effort
        }
        return errorResult("create_reply_draft", err);
      }

      const list = (rs?: GraphRecipient[]) =>
        (rs ?? []).map((r) => r.emailAddress?.address).filter(Boolean).join(", ") || "(none)";
      return {
        content: [
          {
            type: "text",
            text:
              `Reply draft created in Drafts, not sent.\n` +
              `Draft id: ${updated.id}\n` +
              `Subject: ${updated.subject ?? draft.subject ?? ""}\n` +
              `To: ${list(updated.toRecipients)}\n` +
              `Cc: ${list(updated.ccRecipients)}\n` +
              `Conversation id: ${updated.conversationId ?? draft.conversationId ?? ""}` +
              (updated.webLink ?? draft.webLink
                ? `\nOpen: ${updated.webLink ?? draft.webLink}`
                : ""),
          },
        ],
      };
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

// ============================================================================
// Reply-draft helpers
// ============================================================================

interface GraphRecipient {
  emailAddress?: { address?: string; name?: string };
}

interface DraftMessage {
  id: string;
  subject?: string;
  conversationId?: string;
  webLink?: string;
  body?: { contentType?: string; content?: string };
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
}

function escapeHtml(s: string) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Plain text to simple HTML: paragraphs on blank lines, <br> on single breaks. */
export function textToHtml(text: string) {
  return text
    .replace(/\r\n/g, "\n")
    .split(/\n{2,}/)
    .map((para) => `<p>${escapeHtml(para).replace(/\n/g, "<br>")}</p>`)
    .join("");
}

function htmlToRoughText(html: string) {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<[^>]+>/g, "")
    .trim();
}

/** Put new content directly after the opening <body> tag, above the quoted history. */
export function insertAfterBodyTag(html: string, insert: string) {
  const m = /<body[^>]*>/i.exec(html);
  if (!m) return insert + html;
  const at = m.index + m[0].length;
  return html.slice(0, at) + insert + html.slice(at);
}

/** Existing recipients first, then any new ones not already present. */
export function mergeRecipients(existing: GraphRecipient[], extra: GraphRecipient[]) {
  const seen = new Set(
    existing.map((r) => r.emailAddress?.address?.toLowerCase()).filter(Boolean) as string[]
  );
  const out = [...existing];
  for (const r of extra) {
    const a = r.emailAddress?.address?.toLowerCase();
    if (a && !seen.has(a)) {
      seen.add(a);
      out.push(r);
    }
  }
  return out;
}
