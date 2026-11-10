/**
 * Outlook Contacts tools.
 *
 * Contacts are matched by email address, so callers (a daily sync, a backfill)
 * can push the same person repeatedly without creating duplicates. Only the
 * fields a caller sends are written; omitted fields are left as they are in
 * Outlook, so a partial update never blanks out data the user typed by hand.
 *
 * Requires the delegated Graph permission Contacts.ReadWrite on the Entra app
 * (see docs/ENTRA_SETUP.md).
 */
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graph, GraphError } from "../graph";

const MAX_BATCH = 50;

const addressSchema = z.object({
  street: z.string().optional(),
  city: z.string().optional(),
  state: z.string().optional(),
  postal_code: z.string().optional(),
  country: z.string().optional(),
});

const contactSchema = z.object({
  email: z.string().email().describe("Primary email; used to find an existing card"),
  given_name: z.string().optional(),
  surname: z.string().optional(),
  display_name: z
    .string()
    .optional()
    .describe("Defaults to 'given_name surname' when either is sent"),
  company: z.string().optional(),
  job_title: z.string().optional(),
  business_phones: z
    .array(z.string())
    .max(2)
    .optional()
    .describe("Work line first, then main line. Replaces the card's business phones."),
  mobile_phone: z.string().optional(),
  business_address: addressSchema.optional(),
  categories: z
    .array(z.string())
    .optional()
    .describe("Outlook categories to ADD to the card (existing ones are kept)"),
  notes: z.string().optional().describe("Replaces the card's notes"),
});

type ContactInput = z.infer<typeof contactSchema>;

interface GraphContact {
  id: string;
  displayName?: string;
  categories?: string[];
  emailAddresses?: Array<{ address?: string; name?: string }>;
}

/** Escape a value for use inside an OData single-quoted string literal. */
function odataString(value: string): string {
  return value.replace(/'/g, "''");
}

async function findByEmail(session: string, email: string): Promise<GraphContact[]> {
  const filter = `emailAddresses/any(a:a/address eq '${odataString(email)}')`;
  const path =
    `/me/contacts?$filter=${encodeURIComponent(filter)}` +
    `&$select=id,displayName,categories,emailAddresses&$top=10`;
  const res = (await graph(session, path)) as { value?: GraphContact[] };
  return res.value ?? [];
}

function toGraphBody(c: ContactInput, existing?: GraphContact): Record<string, unknown> {
  const body: Record<string, unknown> = {};
  if (c.given_name !== undefined) body.givenName = c.given_name;
  if (c.surname !== undefined) body.surname = c.surname;
  const fullName = [c.given_name, c.surname].filter(Boolean).join(" ").trim();
  const displayName = c.display_name ?? (fullName || undefined);
  if (displayName !== undefined) body.displayName = displayName;
  if (c.company !== undefined) body.companyName = c.company;
  if (c.job_title !== undefined) body.jobTitle = c.job_title;
  if (c.business_phones !== undefined) body.businessPhones = c.business_phones.filter(Boolean);
  if (c.mobile_phone !== undefined) body.mobilePhone = c.mobile_phone;
  if (c.notes !== undefined) body.personalNotes = c.notes;
  if (c.business_address) {
    const a = c.business_address;
    body.businessAddress = {
      street: a.street ?? "",
      city: a.city ?? "",
      state: a.state ?? "",
      postalCode: a.postal_code ?? "",
      countryOrRegion: a.country ?? "",
    };
  }
  if (c.categories && c.categories.length > 0) {
    const merged = new Set([...(existing?.categories ?? []), ...c.categories]);
    body.categories = [...merged];
  }

  // Email: keep the card's other addresses, make sure this one is present and
  // carries the display name Outlook shows in autocomplete.
  const others = (existing?.emailAddresses ?? []).filter(
    (e) => e.address && e.address.toLowerCase() !== c.email.toLowerCase()
  );
  body.emailAddresses = [
    { address: c.email, name: displayName ?? existing?.displayName ?? c.email },
    ...others,
  ].slice(0, 3); // Graph allows at most 3 email addresses per contact
  return body;
}

type UpsertOutcome =
  | { email: string; action: "created" | "updated"; id: string; duplicates?: number }
  | { email: string; action: "error"; error: string };

async function upsertOne(session: string, c: ContactInput): Promise<UpsertOutcome> {
  try {
    const email = c.email.trim();
    let matches = await findByEmail(session, email);
    if (matches.length === 0 && email !== email.toLowerCase()) {
      matches = await findByEmail(session, email.toLowerCase());
    }
    if (matches.length === 0) {
      const created = (await graph(session, "/me/contacts", {
        method: "POST",
        body: JSON.stringify(toGraphBody({ ...c, email })),
      })) as { id: string };
      return { email, action: "created", id: created.id };
    }
    const target = matches[0];
    await graph(session, `/me/contacts/${encodeURIComponent(target.id)}`, {
      method: "PATCH",
      body: JSON.stringify(toGraphBody({ ...c, email }, target)),
    });
    return {
      email,
      action: "updated",
      id: target.id,
      ...(matches.length > 1 ? { duplicates: matches.length - 1 } : {}),
    };
  } catch (err) {
    const msg =
      err instanceof GraphError
        ? `${err.message} (HTTP ${err.statusCode}${err.graphCode ? `, ${err.graphCode}` : ""})`
        : err instanceof Error
          ? err.message
          : String(err);
    return { email: c.email, action: "error", error: msg };
  }
}

export function registerContactTools(server: McpServer) {
  // ==========================================================================
  // upsert_contacts
  // ==========================================================================
  server.tool(
    "upsert_contacts",
    "Create or update Outlook contact cards in the user's default Contacts folder, " +
      `matched by email address (1–${MAX_BATCH} per call). Only the fields you send ` +
      "are written; omitted fields are left unchanged on existing cards. Categories " +
      "are added, never removed. Returns one line per contact: created, updated, or error.",
    {
      contacts: z.array(contactSchema).min(1).max(MAX_BATCH),
    },
    async ({ contacts }, { authInfo }) => {
      const session = authInfo!.extra!.accessToken as string;
      const results: UpsertOutcome[] = [];
      // Sequential on purpose: Graph throttles bursts of writes to one mailbox.
      for (const c of contacts) {
        results.push(await upsertOne(session, c));
      }
      const count = (a: UpsertOutcome["action"]) => results.filter((r) => r.action === a).length;
      const lines = results.map((r) =>
        r.action === "error"
          ? `error   ${r.email}: ${r.error}`
          : `${r.action} ${r.email}` +
            ("duplicates" in r && r.duplicates
              ? ` (note: ${r.duplicates} other card(s) share this email)`
              : "")
      );
      return {
        content: [
          {
            type: "text",
            text:
              `Created ${count("created")}, updated ${count("updated")}, ` +
              `errors ${count("error")}.\n` +
              lines.join("\n"),
          },
        ],
        isError: count("error") === results.length,
      };
    }
  );

  // ==========================================================================
  // find_contacts
  // ==========================================================================
  server.tool(
    "find_contacts",
    "Look up Outlook contact cards by exact email address, or list cards in an " +
      "Outlook category. Returns name, company, title, phones, address and id.",
    {
      email: z.string().email().optional(),
      category: z.string().optional().describe("e.g. 'Wired CIO Deals'"),
      limit: z.number().int().min(1).max(200).default(50),
    },
    async ({ email, category, limit }, { authInfo }) => {
      const session = authInfo!.extra!.accessToken as string;
      if (!email && !category) {
        return {
          content: [{ type: "text", text: "Send an email or a category." }],
          isError: true,
        };
      }
      const filter = email
        ? `emailAddresses/any(a:a/address eq '${odataString(email)}')`
        : `categories/any(c:c eq '${odataString(category!)}')`;
      const select =
        "id,displayName,givenName,surname,companyName,jobTitle,businessPhones," +
        "mobilePhone,businessAddress,emailAddresses,categories";
      try {
        const res = (await graph(
          session,
          `/me/contacts?$filter=${encodeURIComponent(filter)}&$select=${select}&$top=${limit}`
        )) as { value?: unknown[] };
        const items = res.value ?? [];
        return {
          content: [
            {
              type: "text",
              text: items.length
                ? JSON.stringify(items, null, 1)
                : "No matching contact cards.",
            },
          ],
        };
      } catch (err) {
        const msg =
          err instanceof GraphError
            ? `Microsoft Graph error in find_contacts: ${err.message}\nHTTP ${err.statusCode}` +
              (err.graphCode ? ` (${err.graphCode})` : "")
            : `Error in find_contacts: ${err instanceof Error ? err.message : String(err)}`;
        return { content: [{ type: "text", text: msg }], isError: true };
      }
    }
  );
}
