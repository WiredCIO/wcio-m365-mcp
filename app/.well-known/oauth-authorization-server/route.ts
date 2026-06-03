/**
 * OAuth 2.0 Authorization Server Metadata (RFC 8414).
 *
 * Advertises this deployment as the authorization server for the MCP resource.
 * Toward downstream clients (Claude) we are a PKCE-only AS with no client
 * authentication (`token_endpoint_auth_methods: ["none"]`) that supports
 * dynamic client registration — which Microsoft Entra does not, and which MCP
 * clients require. (We are separately a *confidential* client toward Entra; see
 * lib/entra.ts — that relationship is invisible to downstream clients.)
 */
import { NextResponse } from "next/server";
import { BASE_URL } from "@/lib/env";

// Dynamic (not force-static): env validation runs at module load, so
// prerendering this at build time would require runtime secrets during the
// build. Served per-request instead; still cached via Cache-Control.
export const dynamic = "force-dynamic";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "*",
};

export async function GET() {
  return NextResponse.json(
    {
      issuer: BASE_URL,
      authorization_endpoint: `${BASE_URL}/api/oauth/authorize`,
      token_endpoint: `${BASE_URL}/api/oauth/token`,
      registration_endpoint: `${BASE_URL}/api/oauth/register`,
      response_types_supported: ["code"],
      response_modes_supported: ["query"],
      grant_types_supported: ["authorization_code", "refresh_token"],
      token_endpoint_auth_methods_supported: ["none"],
      code_challenge_methods_supported: ["S256"],
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
