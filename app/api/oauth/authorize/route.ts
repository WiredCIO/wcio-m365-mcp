/**
 * Authorization endpoint.
 *
 * Validates the MCP client's PKCE authorization request, stashes it, then
 * redirects the user's browser to Microsoft Entra to actually authenticate.
 * Crucially, the request we send to Entra carries only `scope` (no `resource`
 * parameter) so Entra v2.0 doesn't reject it with AADSTS9010010.
 */
import { NextResponse } from "next/server";
import {
  ENTRA_AUTHORIZE_ENDPOINT,
  ENTRA_REDIRECT_URI,
  GRAPH_SCOPES,
  env,
} from "@/lib/env";
import { generateToken, generatePkce, isRegisteredRedirectUri, buildClientRedirect } from "@/lib/oauth";
import { getClient, putAuthRequest } from "@/lib/store";

export const dynamic = "force-dynamic";

function errorPage(message: string, status = 400) {
  return new NextResponse(`Authorization error: ${message}`, {
    status,
    headers: { "Content-Type": "text/plain" },
  });
}

export async function GET(req: Request) {
  const url = new URL(req.url);
  const p = url.searchParams;

  const clientId = p.get("client_id") ?? "";
  const redirectUri = p.get("redirect_uri") ?? "";
  const responseType = p.get("response_type") ?? "";
  const codeChallenge = p.get("code_challenge") ?? "";
  const codeChallengeMethod = p.get("code_challenge_method") ?? "";
  const clientState = p.get("state") ?? undefined;

  // 1. Validate the client and its redirect URI BEFORE trusting the redirect.
  const client = await getClient(clientId);
  if (!client) {
    return errorPage("unknown client_id");
  }
  if (!redirectUri || !isRegisteredRedirectUri(redirectUri, client.redirectUris)) {
    return errorPage("redirect_uri does not match a registered value");
  }

  // 2. Validate the rest of the request. Now that redirect_uri is trusted, we
  //    can report protocol errors back to the client per OAuth.
  const fail = (error: string, desc: string) =>
    NextResponse.redirect(
      buildClientRedirect(redirectUri, {
        error,
        error_description: desc,
        ...(clientState ? { state: clientState } : {}),
      }),
      302
    );

  if (responseType !== "code") {
    return fail("unsupported_response_type", "only response_type=code is supported");
  }
  if (!codeChallenge || codeChallengeMethod !== "S256") {
    return fail("invalid_request", "PKCE with code_challenge_method=S256 is required");
  }

  // 3. Stash the request and start our own (separately PKCE-protected) leg to Entra.
  const state = generateToken(24);
  const entra = generatePkce();
  await putAuthRequest(state, {
    clientId,
    clientRedirectUri: redirectUri,
    clientState,
    clientCodeChallenge: codeChallenge,
    clientCodeChallengeMethod: codeChallengeMethod,
    entraVerifier: entra.verifier,
    createdAt: Date.now(),
  });

  const entraUrl = new URL(ENTRA_AUTHORIZE_ENDPOINT);
  entraUrl.searchParams.set("client_id", env.WCIO_MCP_CLIENT_ID);
  entraUrl.searchParams.set("response_type", "code");
  entraUrl.searchParams.set("redirect_uri", ENTRA_REDIRECT_URI);
  entraUrl.searchParams.set("response_mode", "query");
  entraUrl.searchParams.set("scope", GRAPH_SCOPES);
  entraUrl.searchParams.set("state", state);
  entraUrl.searchParams.set("code_challenge", entra.challenge);
  entraUrl.searchParams.set("code_challenge_method", "S256");
  // Note: deliberately NO `resource` parameter (see module header).

  return NextResponse.redirect(entraUrl.toString(), 302);
}
