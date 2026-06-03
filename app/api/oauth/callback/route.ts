/**
 * Entra redirect callback.
 *
 * Entra sends the user back here with an authorization code. We exchange it
 * server-side (confidential client) for Graph access + refresh tokens, create
 * a session, then mint OUR own one-time code and bounce the browser back to
 * the MCP client's redirect URI.
 */
import { NextResponse } from "next/server";
import { exchangeCodeForTokens } from "@/lib/entra";
import { buildClientRedirect, generateToken } from "@/lib/oauth";
import {
  encrypt,
  putAuthCode,
  putSession,
  takeAuthRequest,
} from "@/lib/store";
import { VERBOSE } from "@/lib/env";

export const dynamic = "force-dynamic";

function errorPage(message: string, status = 400) {
  return new NextResponse(`Sign-in error: ${message}`, {
    status,
    headers: { "Content-Type": "text/plain" },
  });
}

export async function GET(req: Request) {
  const p = new URL(req.url).searchParams;
  const state = p.get("state") ?? "";
  const code = p.get("code");
  const entraError = p.get("error");
  const entraErrorDesc = p.get("error_description") ?? "";

  // Look up (and consume) the in-flight request keyed by our state.
  const authReq = await takeAuthRequest(state);
  if (!authReq) {
    return errorPage("invalid or expired authorization state");
  }

  // Entra reported an error — relay it to the client.
  if (entraError) {
    return NextResponse.redirect(
      buildClientRedirect(authReq.clientRedirectUri, {
        error: entraError,
        error_description: entraErrorDesc,
        ...(authReq.clientState ? { state: authReq.clientState } : {}),
      }),
      302
    );
  }

  if (!code) {
    return errorPage("missing authorization code");
  }

  // Exchange the Entra code for Graph tokens (server-side, with client secret).
  let tokens;
  try {
    tokens = await exchangeCodeForTokens(code, authReq.entraVerifier);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (VERBOSE) console.error(`[callback] token exchange failed: ${msg}`);
    return NextResponse.redirect(
      buildClientRedirect(authReq.clientRedirectUri, {
        error: "server_error",
        error_description: "Token exchange with the identity provider failed",
        ...(authReq.clientState ? { state: authReq.clientState } : {}),
      }),
      302
    );
  }

  // Persist the session with encrypted Entra tokens.
  const sessionId = generateToken(24);
  await putSession({
    sessionId,
    userOid: tokens.userOid,
    userIdentifier: tokens.userIdentifier,
    graphAccessTokenEnc: encrypt(tokens.accessToken),
    graphRefreshTokenEnc: encrypt(tokens.refreshToken),
    graphExpiresAt: Date.now() + tokens.expiresIn * 1000,
    scopes: tokens.scope ? tokens.scope.split(" ").filter(Boolean) : [],
    createdAt: Date.now(),
  });

  // Mint our own one-time authorization code bound to the client's PKCE.
  const ourCode = generateToken(24);
  await putAuthCode(ourCode, {
    clientId: authReq.clientId,
    redirectUri: authReq.clientRedirectUri,
    clientCodeChallenge: authReq.clientCodeChallenge,
    clientCodeChallengeMethod: authReq.clientCodeChallengeMethod,
    sessionId,
  });

  return NextResponse.redirect(
    buildClientRedirect(authReq.clientRedirectUri, {
      code: ourCode,
      ...(authReq.clientState ? { state: authReq.clientState } : {}),
    }),
    302
  );
}
