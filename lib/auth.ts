/**
 * Validates Microsoft Entra ID access tokens.
 *
 * We expect tokens issued by Microsoft Entra (Azure AD) for our specific
 * tenant, with an audience of Microsoft Graph. The signature is checked
 * against Microsoft's published JWKS endpoint (cached and rotated by `jose`).
 *
 * On success we return AuthInfo with the original raw token in `extra.accessToken`
 * so individual tool handlers can pass it through to Graph.
 */
import { createRemoteJWKSet, jwtVerify, type JWTPayload } from "jose";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import { env, ISSUER, GRAPH_AUDIENCE, VERBOSE } from "./env";

/**
 * Microsoft's tenant-scoped JWKS endpoint. `jose` caches keys and respects
 * cache-control headers — no manual TTL needed.
 */
const JWKS = createRemoteJWKSet(
  new URL(`https://login.microsoftonline.com/${env.WCIO_MCP_TENANT_ID}/discovery/v2.0/keys`)
);

interface EntraClaims extends JWTPayload {
  /** Object ID of the user — stable identifier within the tenant. */
  oid?: string;
  /** Tenant ID — must match our configured tenant. */
  tid?: string;
  /** Space-separated delegated permissions (Graph scopes). */
  scp?: string;
  /** User's UPN / email for logging. */
  upn?: string;
  preferred_username?: string;
  /** App ID that requested the token. */
  appid?: string;
  azp?: string;
}

/**
 * Verify a Bearer token and return AuthInfo for the MCP request.
 * Returns `undefined` for invalid tokens, which causes mcp-handler to
 * respond with a 401 and the WWW-Authenticate header.
 */
export async function verifyEntraToken(
  _req: Request,
  bearerToken?: string
): Promise<AuthInfo | undefined> {
  if (!bearerToken) return undefined;

  try {
    const { payload } = await jwtVerify<EntraClaims>(bearerToken, JWKS, {
      issuer: [
        ISSUER,
        // v1.0 issuer format — Microsoft sometimes mixes them
        `https://sts.windows.net/${env.WCIO_MCP_TENANT_ID}/`,
      ],
      audience: GRAPH_AUDIENCE,
    });

    // Defense in depth: explicitly check tenant ID even though the issuer
    // check above already pins us to it.
    if (payload.tid !== env.WCIO_MCP_TENANT_ID) {
      if (VERBOSE) {
        console.warn(`[auth] rejected: tid mismatch (got ${payload.tid})`);
      }
      return undefined;
    }

    const scopes = (payload.scp ?? "").split(" ").filter(Boolean);
    const userIdentifier =
      payload.upn ?? payload.preferred_username ?? payload.oid ?? "unknown";

    if (VERBOSE) {
      console.log(
        `[auth] accepted token for ${userIdentifier} (scopes: ${scopes.join(",")})`
      );
    }

    return {
      token: bearerToken,
      clientId: payload.appid ?? payload.azp ?? "unknown",
      scopes,
      extra: {
        // The raw token is what we forward to Graph for tool calls.
        accessToken: bearerToken,
        userId: payload.oid,
        userIdentifier,
        tenantId: payload.tid,
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
 * Convenience: extract the access token from AuthInfo for a tool handler.
 */
export function getAccessToken(authInfo: AuthInfo | undefined): string {
  const token = authInfo?.extra?.accessToken;
  if (typeof token !== "string") {
    throw new Error("No access token available — authentication required");
  }
  return token;
}

/**
 * Convenience: extract the user identifier (UPN/email) for audit logging.
 */
export function getUserIdentifier(authInfo: AuthInfo | undefined): string {
  const id = authInfo?.extra?.userIdentifier;
  return typeof id === "string" ? id : "unknown";
}
