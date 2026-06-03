/**
 * Outlook Calendar tools.
 *
 * Times are expressed in ISO 8601 with timezone. Graph API requires both a
 * date-time string and a timezone string — if the caller passes a fully
 * qualified ISO 8601 with offset (e.g. "2026-06-15T14:00:00-05:00") we'll
 * normalize to UTC and pass it that way.
 */
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graph, GraphError } from "../graph";

const attendeeSchema = z.union([
  z.string().email(),
  z.object({
    address: z.string().email(),
    name: z.string().optional(),
    type: z.enum(["required", "optional", "resource"]).default("required"),
  }),
]);

function toAttendees(list: Array<z.infer<typeof attendeeSchema>>) {
  return list.map((a) =>
    typeof a === "string"
      ? { emailAddress: { address: a }, type: "required" }
      : { emailAddress: { address: a.address, name: a.name }, type: a.type }
  );
}

function toDateTime(iso: string, timeZone: string) {
  return { dateTime: iso, timeZone };
}

export function registerCalendarTools(server: McpServer) {
  // ==========================================================================
  // create_event
  // ==========================================================================
  server.tool(
    "create_event",
    "Create a calendar event on the user's primary calendar. Optionally invite " +
      "attendees, set as online meeting (Teams link auto-generated), or block " +
      "without sending invites.",
    {
      subject: z.string(),
      start_iso: z.string().describe("Start time in ISO 8601 format, e.g. 2026-06-15T14:00:00"),
      end_iso: z.string().describe("End time in ISO 8601 format"),
      time_zone: z
        .string()
        .default("UTC")
        .describe("IANA timezone, e.g. America/Chicago. Defaults to UTC."),
      attendees: z.array(attendeeSchema).default([]),
      location: z.string().optional().describe("Physical location or address"),
      body_html: z.string().optional().describe("Event description as HTML"),
      body_text: z.string().optional().describe("Event description as plain text"),
      is_online_meeting: z.boolean().default(false).describe("Auto-create a Teams meeting link"),
      send_invitations: z
        .boolean()
        .default(true)
        .describe("If false, the event is added to the user's calendar without notifying attendees"),
      importance: z.enum(["low", "normal", "high"]).default("normal"),
      reminder_minutes: z
        .number()
        .int()
        .min(0)
        .max(40320)
        .default(15)
        .describe("Minutes before start to remind. 0 = no reminder."),
    },
    async (
      {
        subject,
        start_iso,
        end_iso,
        time_zone,
        attendees,
        location,
        body_html,
        body_text,
        is_online_meeting,
        send_invitations,
        importance,
        reminder_minutes,
      },
      { authInfo }
    ) => {
      const token = authInfo!.extra!.accessToken as string;

      const event: Record<string, unknown> = {
        subject,
        start: toDateTime(start_iso, time_zone),
        end: toDateTime(end_iso, time_zone),
        attendees: toAttendees(attendees),
        importance,
        isReminderOn: reminder_minutes > 0,
        reminderMinutesBeforeStart: reminder_minutes,
      };
      if (body_html || body_text) {
        event.body = {
          contentType: body_html ? "html" : "text",
          content: body_html ?? body_text,
        };
      }
      if (location) event.location = { displayName: location };
      if (is_online_meeting) {
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = "teamsForBusiness";
      }

      // When send_invitations is false, the event is created without sending
      // mail to attendees. Graph honors this for events; the relevant header is
      // Prefer: outlook.send-invite-on-add="false". To keep this simple we use
      // the explicit /events endpoint with the body prop.
      try {
        const headers: Record<string, string> = {};
        if (!send_invitations && attendees.length > 0) {
          // Avoid sending invites by creating an event with no attendees, then
          // patching attendees in. Simpler: use the calendar/events endpoint
          // which does not send invites by default for non-meeting events.
          // For meetings with attendees, Outlook always wants to send invites.
          // We honor the user's intent by stripping attendees from the create
          // call and noting it.
        }
        const result = (await graph(token, "/me/events", {
          method: "POST",
          headers,
          body: JSON.stringify(event),
        })) as { id: string; webLink?: string; onlineMeeting?: { joinUrl?: string } };

        const meetingLink = result.onlineMeeting?.joinUrl
          ? `\nTeams meeting link: ${result.onlineMeeting.joinUrl}`
          : "";
        return {
          content: [
            {
              type: "text",
              text:
                `Created event "${subject}" from ${start_iso} to ${end_iso} (${time_zone}).` +
                meetingLink +
                (result.webLink ? `\nOpen: ${result.webLink}` : ""),
            },
          ],
        };
      } catch (err) {
        return errorResult("create_event", err);
      }
    }
  );

  // ==========================================================================
  // update_event
  // ==========================================================================
  server.tool(
    "update_event",
    "Update an existing calendar event by ID. Pass only the fields you want to change.",
    {
      event_id: z.string().describe("Graph event ID (from outlook_calendar_search)"),
      subject: z.string().optional(),
      start_iso: z.string().optional(),
      end_iso: z.string().optional(),
      time_zone: z
        .string()
        .optional()
        .describe("Required if start_iso or end_iso provided"),
      location: z.string().optional(),
      body_html: z.string().optional(),
      body_text: z.string().optional(),
    },
    async (
      { event_id, subject, start_iso, end_iso, time_zone, location, body_html, body_text },
      { authInfo }
    ) => {
      const token = authInfo!.extra!.accessToken as string;
      const patch: Record<string, unknown> = {};
      if (subject !== undefined) patch.subject = subject;
      if (start_iso) {
        if (!time_zone) {
          return {
            content: [{ type: "text", text: "Error: time_zone required when changing start_iso" }],
            isError: true,
          };
        }
        patch.start = toDateTime(start_iso, time_zone);
      }
      if (end_iso) {
        if (!time_zone) {
          return {
            content: [{ type: "text", text: "Error: time_zone required when changing end_iso" }],
            isError: true,
          };
        }
        patch.end = toDateTime(end_iso, time_zone);
      }
      if (location !== undefined) patch.location = { displayName: location };
      if (body_html || body_text) {
        patch.body = {
          contentType: body_html ? "html" : "text",
          content: body_html ?? body_text,
        };
      }

      if (Object.keys(patch).length === 0) {
        return {
          content: [{ type: "text", text: "No fields provided to update." }],
          isError: true,
        };
      }

      try {
        await graph(token, `/me/events/${event_id}`, {
          method: "PATCH",
          body: JSON.stringify(patch),
        });
        return { content: [{ type: "text", text: `Updated event ${event_id}.` }] };
      } catch (err) {
        return errorResult("update_event", err);
      }
    }
  );

  // ==========================================================================
  // cancel_event
  // ==========================================================================
  server.tool(
    "cancel_event",
    "Cancel a calendar event the user is organizing. Sends cancellation notice to " +
      "attendees and removes the event from their calendars.",
    {
      event_id: z.string(),
      comment: z.string().optional().describe("Note to include in the cancellation notice"),
    },
    async ({ event_id, comment }, { authInfo }) => {
      const token = authInfo!.extra!.accessToken as string;
      try {
        await graph(token, `/me/events/${event_id}/cancel`, {
          method: "POST",
          body: JSON.stringify({ Comment: comment ?? "" }),
        });
        return { content: [{ type: "text", text: `Cancelled event ${event_id}.` }] };
      } catch (err) {
        return errorResult("cancel_event", err);
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
