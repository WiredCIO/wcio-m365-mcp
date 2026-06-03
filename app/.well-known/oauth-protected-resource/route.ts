/**
 * OAuth 2.0 Protected Resource Metadata (RFC 9728).
 *
 * Points MCP clients at OUR authorization server (this same deployment), which
 * proxies Microsoft Entra server-side. We do NOT point clients directly at
 * Entra: Entra's v2.0 endpoint rejects the `resource` + v2 `scope` combination
 * that MCP clients send (AADSTS9010010), and Entra has no dynamic client
 * registration. Fronting it with our own AS solves both.
 */
import { NextResponse } from "next/server";
import { BASE_URL, MCP_RESOURCE } from "@/lib/env";

// Dynamic (not force-static): env validation runs at module load, so
// prerendering this at build time would require runtime secrets to be present
// during the build. Served per-request instead — the response is still cached
// via Cache-Control. The handler only reads BASE_URL/MCP_RESOURCE.
export const dynamic = "force-dynamic";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "*",
};

export async function GET() {
  return NextResponse.json(
    {
      resource: MCP_RESOURCE,
      authorization_servers: [BASE_URL],
      bearer_methods_supported: ["header"],
      scopes_supported: ["mcp"],
    },
    { headers: { ...CORS, "Cache-Control": "public, max-age=300" } }
  );
}

export async function OPTIONS() {
  return new Response(null, {
    status: 204,
    headers: { ...CORS, "Access-Control-Max-Age": "86400" },
  });
}
