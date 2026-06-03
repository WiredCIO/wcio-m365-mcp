/**
 * Microsoft Graph helpers — thin wrapper around fetch that adds the bearer
 * token, sets common headers, and surfaces Graph errors with useful messages.
 *
 * We deliberately don't use the official @microsoft/microsoft-graph-client
 * package for most operations because:
 *  1. It's heavyweight relative to what we need
 *  2. Its auth provider pattern doesn't compose well with per-request tokens
 *  3. Direct fetch gives us better control over streaming for large uploads
 */
import { GRAPH_BASE, TOKEN_ENDPOINT, VERBOSE, env } from "./env";
import { createHash } from "crypto";

export class GraphError extends Error {
  constructor(
    message: string,
    public statusCode: number,
    public graphCode?: string,
    public requestId?: string
  ) {
    super(message);
    this.name = "GraphError";
  }
}

// ============================================================================
// On-Behalf-Of (OBO) token exchange
// ============================================================================
// We receive a Bearer token scoped to our app (audience = our client ID).
// To call Microsoft Graph, we need a Graph-audience token. The OBO grant
// type exchanges the user's app-scoped token for a Graph-scoped token,
// preserving the user's identity throughout.
//
// Cache strategy: in-memory keyed by SHA-256 of the incoming user token.
// Vercel warm functions reuse this map across requests; cold starts lose
// it and we do a fresh exchange. Tokens are cached until ~5 minutes before
// their declared expiry to avoid races.

interface CachedGraphToken {
  graphToken: string;
  expiresAt: number; // epoch millis
}

const graphTokenCache = new Map<string, CachedGraphToken>();
const TOKEN_REFRESH_BUFFER_MS = 5 * 60 * 1000;

function cacheKey(userToken: string): string {
  return createHash("sha256").update(userToken).digest("hex");
}

/**
 * Exchange the user's app-scoped token for a Microsoft Graph-scoped token
 * using OAuth 2.0 On-Behalf-Of grant. Cached per user-token for ~55 minutes.
 *
 * Requests Graph's `.default` scope, which yields a token containing every
 * delegated permission our app registration has admin-consented for the
 * user — so we don't need per-tool scope arguments.
 *
 * https://learn.microsoft.com/entra/identity-platform/v2-oauth2-on-behalf-of-flow
 */
export async function exchangeForGraphToken(userToken: string): Promise<string> {
  const key = cacheKey(userToken);
  const cached = graphTokenCache.get(key);
  const now = Date.now();
  if (cached && cached.expiresAt - TOKEN_REFRESH_BUFFER_MS > now) {
    return cached.graphToken;
  }

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    client_id: env.WCIO_MCP_CLIENT_ID,
    client_secret: env.WCIO_MCP_CLIENT_SECRET,
    assertion: userToken,
    scope: "https://graph.microsoft.com/.default offline_access",
    requested_token_use: "on_behalf_of",
  });

  const res = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    let detail = "";
    try {
      const errJson = (await res.json()) as { error?: string; error_description?: string };
      detail = `${errJson.error ?? "unknown_error"}: ${errJson.error_description ?? ""}`;
    } catch {
      detail = await res.text();
    }
    throw new GraphError(
      `OBO token exchange failed: ${detail}`,
      res.status,
      "obo_exchange_failed"
    );
  }

  const tokenResponse = (await res.json()) as {
    access_token: string;
    expires_in: number;
    token_type: string;
  };

  graphTokenCache.set(key, {
    graphToken: tokenResponse.access_token,
    expiresAt: now + tokenResponse.expires_in * 1000,
  });

  if (VERBOSE) {
    console.log(`[obo] exchanged user token for Graph token, valid ${tokenResponse.expires_in}s`);
  }

  // Periodic cleanup so the cache doesn't grow unboundedly on a warm instance
  if (graphTokenCache.size > 200) {
    for (const [k, v] of graphTokenCache.entries()) {
      if (v.expiresAt < now) graphTokenCache.delete(k);
    }
  }

  return tokenResponse.access_token;
}

interface GraphErrorBody {
  error?: {
    code?: string;
    message?: string;
    innerError?: {
      "request-id"?: string;
      date?: string;
    };
  };
}

/**
 * Make a Graph API request. Accepts the user's MCP-scoped Bearer token;
 * internally exchanges it for a Graph-scoped token via OBO before calling.
 * Returns parsed JSON on success. Throws GraphError with HTTP status on failure.
 */
export async function graph<T = unknown>(
  userToken: string,
  path: string,
  init: RequestInit = {}
): Promise<T> {
  const graphToken = await exchangeForGraphToken(userToken);
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${graphToken}`);
  if (!headers.has("Content-Type") && init.body && typeof init.body === "string") {
    headers.set("Content-Type", "application/json");
  }
  if (!headers.has("Accept")) {
    headers.set("Accept", "application/json");
  }

  const response = await fetch(url, { ...init, headers });

  if (VERBOSE) {
    console.log(`[graph] ${init.method ?? "GET"} ${path} → ${response.status}`);
  }

  if (!response.ok) {
    let errBody: GraphErrorBody | undefined;
    try {
      errBody = (await response.json()) as GraphErrorBody;
    } catch {
      // body wasn't JSON
    }
    throw new GraphError(
      errBody?.error?.message ?? `Graph API returned ${response.status}`,
      response.status,
      errBody?.error?.code,
      errBody?.error?.innerError?.["request-id"]
    );
  }

  if (response.status === 204) {
    return undefined as T;
  }

  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return (await response.json()) as T;
  }
  return undefined as T;
}

/**
 * Upload binary content to Graph. Used for file uploads under 4 MB.
 * Takes the user's MCP-scoped token; internally exchanges to Graph.
 */
export async function graphPutBinary(
  userToken: string,
  path: string,
  body: ArrayBuffer | Uint8Array | Blob,
  contentType: string = "application/octet-stream"
): Promise<unknown> {
  const graphToken = await exchangeForGraphToken(userToken);
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${graphToken}`,
      "Content-Type": contentType,
    },
    body: body as BodyInit,
  });

  if (!res.ok) {
    let errBody: GraphErrorBody | undefined;
    try {
      errBody = (await res.json()) as GraphErrorBody;
    } catch {
      // not JSON
    }
    throw new GraphError(
      errBody?.error?.message ?? `Graph PUT returned ${res.status}`,
      res.status,
      errBody?.error?.code,
      errBody?.error?.innerError?.["request-id"]
    );
  }
  if (res.status === 204) return undefined;
  return await res.json();
}

/**
 * Upload a file >4 MB using an upload session.
 * Takes the user's MCP-scoped token; exchanges via OBO for the session-create
 * call. The actual chunked PUTs use the temporary uploadUrl Microsoft returns,
 * which is pre-authenticated so no token needed on those.
 *
 * https://learn.microsoft.com/graph/api/driveitem-createuploadsession
 */
export async function uploadFileViaSession(
  userToken: string,
  uploadSessionPath: string,
  body: ArrayBuffer,
  conflictBehavior: "rename" | "replace" | "fail" = "replace"
): Promise<unknown> {
  // 1. Create the upload session (needs Graph token via OBO)
  const session = (await graph<{ uploadUrl: string }>(
    userToken,
    uploadSessionPath,
    {
      method: "POST",
      body: JSON.stringify({
        item: {
          "@microsoft.graph.conflictBehavior": conflictBehavior,
        },
      }),
    }
  )) as { uploadUrl: string };

  // 2. Stream the bytes in chunks
  const CHUNK_SIZE = 60 * 1024 * 1024; // 60 MB — Graph recommends a multiple of 320 KiB up to 60 MB
  const totalSize = body.byteLength;
  let lastResult: unknown = undefined;

  for (let offset = 0; offset < totalSize; offset += CHUNK_SIZE) {
    const end = Math.min(offset + CHUNK_SIZE, totalSize);
    const chunk = body.slice(offset, end);
    const res = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.byteLength),
        "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
      },
      body: chunk,
    });
    if (!res.ok && res.status !== 202) {
      throw new GraphError(
        `Upload session chunk failed at offset ${offset}`,
        res.status
      );
    }
    if (res.status === 200 || res.status === 201) {
      // Final chunk — Graph returns the DriveItem
      lastResult = await res.json();
    }
  }

  return lastResult;
}

/**
 * Resolve an item path like "/drive/root:/Documents/foo.xlsx" — used to build
 * Graph URLs. Note that for the user's own OneDrive, the base is /me/drive.
 * Group/SharePoint sites use /sites/{siteId}/drive instead.
 */
export function userDrivePath(itemPath: string): string {
  const clean = itemPath.replace(/^\/+/, "");
  return `/me/drive/root:/${clean}`;
}

export function userDriveItemById(itemId: string): string {
  return `/me/drive/items/${itemId}`;
}

export function siteDrivePath(siteId: string, itemPath: string): string {
  const clean = itemPath.replace(/^\/+/, "");
  return `/sites/${siteId}/drive/root:/${clean}`;
}
