/**
 * Validates the opaque Bearer tokens that *we* issued from /api/oauth/token.
 *
 * These are not Entra tokens — they're high-entropy random strings that map
 * (via their SHA-256 hash) to a server-side session holding the user's
 * encrypted Graph tokens. We expose the session id to tool handlers through
 * AuthInfo.extra.accessToken; graph.ts resolves it to a live Graph token.
 */
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import { sha256 } from "./oauth";
import { getAccessTokenRef, getSession } from "./store";
import { VERBOSE } from "./env";

/**
 * Verify a Bearer token and return AuthInfo for the MCP request.
 * Returns `undefined` for invalid/expired tokens, which causes mcp-handler to
 * respond with a 401 and the WWW-Authenticate header.
 */
export async function verifyMcpToken(
  _req: Request,
  bearerToken?: string
): Promise<AuthInfo | undefined> {
  if (!bearerToken) return undefined;

  try {
    const ref = await getAccessTokenRef(sha256(bearerToken));
    if (!ref) {
      if (VERBOSE) console.warn("[auth] rejected: unknown or expired access token");
      return undefined;
    }
    if (ref.expiresAt <= Date.now()) {
      if (VERBOSE) console.warn("[auth] rejected: access token past expiry");
      return undefined;
    }

    const session = await getSession(ref.sessionId);
    if (!session) {
      if (VERBOSE) console.warn("[auth] rejected: session no longer exists");
      return undefined;
    }

    if (VERBOSE) {
      console.log(`[auth] accepted token for ${session.userIdentifier}`);
    }

    return {
      token: bearerToken,
      clientId: ref.clientId,
      scopes: session.scopes,
      extra: {
        // Tool handlers read this as the handle to pass into graph(); it's the
        // opaque session id, not a real Graph token.
        accessToken: session.sessionId,
        sessionId: session.sessionId,
        userId: session.userOid,
        userIdentifier: session.userIdentifier,
      },
    };
  } catch (err) {
    if (VERBOSE) {
      const msg = err instanceof Error ? err.message : String(err);
      console.warn(`[auth] token verification failed: ${msg}`);
    }
    return undefined;
  }
}

/**
 * Convenience: extract the session handle from AuthInfo for a tool handler.
 * (Named getAccessToken for backward compatibility with existing tool code.)
 */
export function getAccessToken(authInfo: AuthInfo | undefined): string {
  const handle = authInfo?.extra?.accessToken;
  if (typeof handle !== "string") {
    throw new Error("No session available — authentication required");
  }
  return handle;
}

/**
 * Convenience: extract the user identifier (UPN/email) for audit logging.
 */
export function getUserIdentifier(authInfo: AuthInfo | undefined): string {
  const id = authInfo?.extra?.userIdentifier;
  return typeof id === "string" ? id : "unknown";
}
