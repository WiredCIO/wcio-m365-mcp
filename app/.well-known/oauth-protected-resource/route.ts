/**
 * OAuth 2.0 Protected Resource Metadata (RFC 9728).
 *
 * Replaces mcp-handler's `protectedResourceHandler` because we need to
 * publish additional fields beyond `authorization_servers`:
 *
 *  - `resource`: our Application ID URI (`api://<client-id>`). This is what
 *    Claude.ai puts in the `resource` parameter when calling Entra's
 *    authorization endpoint. Setting it to our App ID URI makes Entra
 *    happy because the audience of the issued token will match.
 *
 *  - `scopes_supported`: includes our custom scope `api://<client-id>/mcp.access`
 *    so Claude.ai knows what to request. Entra issues a token scoped to
 *    our app with this scope; we then OBO-exchange it for Graph internally.
 *
 *  - `bearer_methods_supported`: ["header"] is the only method we accept.
 */
import { NextResponse } from "next/server";
import { ISSUER, APP_ID_URI, APP_SCOPE } from "@/lib/env";

export const dynamic = "force-static";

export async function GET() {
  return NextResponse.json(
    {
      resource: APP_ID_URI,
      authorization_servers: [ISSUER],
      scopes_supported: [APP_SCOPE, "offline_access", "openid"],
      bearer_methods_supported: ["header"],
    },
    {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, OPTIONS",
        "Access-Control-Allow-Headers": "*",
        "Cache-Control": "public, max-age=300",
      },
    }
  );
}

export async function OPTIONS() {
  return new Response(null, {
    status: 204,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "*",
      "Access-Control-Max-Age": "86400",
    },
  });
}
