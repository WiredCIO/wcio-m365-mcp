/**
 * Token endpoint.
 *
 * Exchanges our one-time authorization code (or a refresh token) for an opaque
 * Bearer access token that the MCP client uses on /api/mcp calls. The access
 * token maps to a server-side session holding the user's encrypted Graph
 * tokens; the client never sees an Entra token.
 */
import { NextResponse } from "next/server";
import { generateToken, sha256, verifyPkceS256 } from "@/lib/oauth";
import {
  getSession,
  putAccessToken,
  putRefreshToken,
  takeAuthCode,
  takeRefreshToken,
  TOKEN_TTL,
} from "@/lib/store";

export const dynamic = "force-dynamic";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

function oauthError(error: string, description: string, status = 400) {
  return NextResponse.json({ error, error_description: description }, { status, headers: CORS });
}

async function readParams(req: Request): Promise<Record<string, string>> {
  const ct = req.headers.get("content-type") ?? "";
  if (ct.includes("application/json")) {
    const j = (await req.json()) as Record<string, unknown>;
    return Object.fromEntries(
      Object.entries(j).map(([k, v]) => [k, typeof v === "string" ? v : String(v)])
    );
  }
  const form = await req.formData();
  const out: Record<string, string> = {};
  for (const [k, v] of form.entries()) out[k] = typeof v === "string" ? v : "";
  return out;
}

/** Issue a fresh access+refresh token pair bound to a session. */
async function issueTokens(sessionId: string, clientId: string) {
  const accessToken = generateToken(32);
  const refreshToken = generateToken(32);
  const expiresAt = Date.now() + TOKEN_TTL.accessToken * 1000;

  await putAccessToken(sha256(accessToken), { sessionId, clientId, expiresAt });
  await putRefreshToken(sha256(refreshToken), { sessionId, clientId, expiresAt });

  return NextResponse.json(
    {
      access_token: accessToken,
      token_type: "Bearer",
      expires_in: TOKEN_TTL.accessToken,
      refresh_token: refreshToken,
      scope: "mcp",
    },
    { headers: { ...CORS, "Cache-Control": "no-store", Pragma: "no-cache" } }
  );
}

export async function POST(req: Request) {
  let params: Record<string, string>;
  try {
    params = await readParams(req);
  } catch {
    return oauthError("invalid_request", "Unparseable request body");
  }

  const grantType = params.grant_type;

  // -------- authorization_code --------
  if (grantType === "authorization_code") {
    const { code, code_verifier, client_id, redirect_uri } = params;
    if (!code || !code_verifier || !client_id) {
      return oauthError("invalid_request", "code, code_verifier and client_id are required");
    }

    const authCode = await takeAuthCode(code);
    if (!authCode) {
      return oauthError("invalid_grant", "authorization code is invalid or expired");
    }
    if (authCode.clientId !== client_id) {
      return oauthError("invalid_grant", "client_id does not match the authorization code");
    }
    if (redirect_uri && redirect_uri !== authCode.redirectUri) {
      return oauthError("invalid_grant", "redirect_uri does not match the authorization request");
    }
    if (!verifyPkceS256(code_verifier, authCode.clientCodeChallenge)) {
      return oauthError("invalid_grant", "PKCE verification failed");
    }

    const session = await getSession(authCode.sessionId);
    if (!session) {
      return oauthError("invalid_grant", "session no longer exists");
    }

    return issueTokens(authCode.sessionId, authCode.clientId);
  }

  // -------- refresh_token --------
  if (grantType === "refresh_token") {
    const { refresh_token, client_id } = params;
    if (!refresh_token) {
      return oauthError("invalid_request", "refresh_token is required");
    }

    const ref = await takeRefreshToken(sha256(refresh_token));
    if (!ref) {
      return oauthError("invalid_grant", "refresh token is invalid or expired");
    }
    if (client_id && ref.clientId !== client_id) {
      return oauthError("invalid_grant", "client_id does not match the refresh token");
    }

    const session = await getSession(ref.sessionId);
    if (!session) {
      return oauthError("invalid_grant", "session no longer exists");
    }

    return issueTokens(ref.sessionId, ref.clientId);
  }

  return oauthError("unsupported_grant_type", `grant_type '${grantType ?? ""}' is not supported`);
}

export async function OPTIONS() {
  return new Response(null, {
    status: 204,
    headers: { ...CORS, "Access-Control-Max-Age": "86400" },
  });
}
