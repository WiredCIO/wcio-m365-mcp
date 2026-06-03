/**
 * Dynamic Client Registration (RFC 7591).
 *
 * MCP clients (Claude) register here to obtain a client_id. We are a public
 * client AS, so no client_secret is issued — clients authenticate with PKCE.
 * Microsoft Entra doesn't offer DCR, which is the whole reason this proxy
 * exists.
 */
import { NextResponse } from "next/server";
import { generateToken, isAcceptableRedirectUri } from "@/lib/oauth";
import { putClient } from "@/lib/store";

export const dynamic = "force-dynamic";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

interface RegistrationBody {
  redirect_uris?: unknown;
  client_name?: unknown;
}

export async function POST(req: Request) {
  let body: RegistrationBody;
  try {
    body = (await req.json()) as RegistrationBody;
  } catch {
    return NextResponse.json(
      { error: "invalid_client_metadata", error_description: "Body must be JSON" },
      { status: 400, headers: CORS }
    );
  }

  const redirectUris = Array.isArray(body.redirect_uris)
    ? body.redirect_uris.filter((u): u is string => typeof u === "string")
    : [];

  if (redirectUris.length === 0) {
    return NextResponse.json(
      { error: "invalid_redirect_uri", error_description: "redirect_uris is required" },
      { status: 400, headers: CORS }
    );
  }

  for (const uri of redirectUris) {
    if (!isAcceptableRedirectUri(uri)) {
      return NextResponse.json(
        {
          error: "invalid_redirect_uri",
          error_description: `Redirect URI not allowed: ${uri} (must be https, or http on loopback)`,
        },
        { status: 400, headers: CORS }
      );
    }
  }

  const clientId = generateToken(16);
  const createdAt = Date.now();
  await putClient({
    clientId,
    redirectUris,
    clientName: typeof body.client_name === "string" ? body.client_name : undefined,
    createdAt,
  });

  return NextResponse.json(
    {
      client_id: clientId,
      redirect_uris: redirectUris,
      token_endpoint_auth_method: "none",
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      client_id_issued_at: Math.floor(createdAt / 1000),
      ...(typeof body.client_name === "string" ? { client_name: body.client_name } : {}),
    },
    { status: 201, headers: CORS }
  );
}

export async function OPTIONS() {
  return new Response(null, {
    status: 204,
    headers: { ...CORS, "Access-Control-Max-Age": "86400" },
  });
}
