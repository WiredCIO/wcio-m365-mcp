/**
 * Server-side Microsoft Entra token exchanges.
 *
 * Both calls happen confidentially (client_id + client_secret) from our
 * backend — Claude never talks to Entra directly. We deliberately never send
 * an OAuth `resource` parameter; the scope already encodes the Graph audience,
 * and Entra v2.0 rejects requests carrying both (AADSTS9010010).
 */
import { decodeJwt } from "jose";
import {
  ENTRA_TOKEN_ENDPOINT,
  ENTRA_REDIRECT_URI,
  GRAPH_SCOPES,
  env,
  VERBOSE,
} from "./env";

export interface EntraTokens {
  accessToken: string;
  refreshToken: string;
  /** Seconds until the access token expires. */
  expiresIn: number;
  scope: string;
  userOid?: string;
  userIdentifier: string;
}

interface EntraTokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
  id_token?: string;
  token_type: string;
}

function identityFromIdToken(idToken?: string): { oid?: string; upn: string } {
  if (!idToken) return { upn: "unknown" };
  try {
    const claims = decodeJwt(idToken) as {
      oid?: string;
      preferred_username?: string;
      upn?: string;
      email?: string;
    };
    return {
      oid: claims.oid,
      upn: claims.preferred_username ?? claims.upn ?? claims.email ?? claims.oid ?? "unknown",
    };
  } catch {
    return { upn: "unknown" };
  }
}

async function postToken(body: URLSearchParams): Promise<EntraTokenResponse> {
  const res = await fetch(ENTRA_TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!res.ok) {
    let detail = "";
    try {
      const j = (await res.json()) as { error?: string; error_description?: string };
      detail = `${j.error ?? "unknown_error"}: ${j.error_description ?? ""}`;
    } catch {
      detail = await res.text();
    }
    throw new Error(`Entra token request failed (${res.status}): ${detail}`);
  }
  return (await res.json()) as EntraTokenResponse;
}

/** Exchange an Entra authorization code (our /callback) for Graph tokens. */
export async function exchangeCodeForTokens(
  code: string,
  codeVerifier: string
): Promise<EntraTokens> {
  const body = new URLSearchParams({
    client_id: env.WCIO_MCP_CLIENT_ID,
    client_secret: env.WCIO_MCP_CLIENT_SECRET,
    grant_type: "authorization_code",
    code,
    redirect_uri: ENTRA_REDIRECT_URI,
    code_verifier: codeVerifier,
    scope: GRAPH_SCOPES,
  });
  const tr = await postToken(body);
  if (!tr.refresh_token) {
    throw new Error("Entra did not return a refresh token (is offline_access consented?)");
  }
  const id = identityFromIdToken(tr.id_token);
  if (VERBOSE) console.log(`[entra] code exchange ok for ${id.upn}, expires ${tr.expires_in}s`);
  return {
    accessToken: tr.access_token,
    refreshToken: tr.refresh_token,
    expiresIn: tr.expires_in,
    scope: tr.scope ?? "",
    userOid: id.oid,
    userIdentifier: id.upn,
  };
}

/** Redeem an Entra refresh token for a fresh Graph access token. */
export async function refreshGraphTokens(refreshToken: string): Promise<EntraTokens> {
  const body = new URLSearchParams({
    client_id: env.WCIO_MCP_CLIENT_ID,
    client_secret: env.WCIO_MCP_CLIENT_SECRET,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: GRAPH_SCOPES,
  });
  const tr = await postToken(body);
  const id = identityFromIdToken(tr.id_token);
  if (VERBOSE) console.log(`[entra] refresh ok, expires ${tr.expires_in}s`);
  return {
    accessToken: tr.access_token,
    // Entra rotates refresh tokens; fall back to the old one if absent.
    refreshToken: tr.refresh_token ?? refreshToken,
    expiresIn: tr.expires_in,
    scope: tr.scope ?? "",
    userOid: id.oid,
    userIdentifier: id.upn,
  };
}
