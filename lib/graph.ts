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
import { GRAPH_BASE, VERBOSE } from "./env";
import { refreshGraphTokens } from "./entra";
import { decrypt, encrypt, getSession, putSession, type Session } from "./store";

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
// Session -> Graph token resolution
// ============================================================================
// Tool handlers pass us an opaque session id (carried through AuthInfo). We
// resolve it to a live Microsoft Graph access token, transparently refreshing
// via the stored Entra refresh token when the cached one is near expiry. The
// OBO flow is gone — the proxy obtained Graph tokens directly during login.

const TOKEN_REFRESH_BUFFER_MS = 5 * 60 * 1000;

/**
 * Resolve a session id to a valid Graph access token. Refreshes through Entra
 * (refresh_token grant) and re-persists the rotated tokens when needed.
 */
export async function getGraphTokenForSession(sessionId: string): Promise<string> {
  const session = await getSession(sessionId);
  if (!session) {
    throw new GraphError("Session not found or expired — re-authentication required", 401);
  }

  const now = Date.now();
  if (session.graphExpiresAt - TOKEN_REFRESH_BUFFER_MS > now) {
    return decrypt(session.graphAccessTokenEnc);
  }

  // Access token expired (or close to it) — refresh via Entra.
  let tokens;
  try {
    tokens = await refreshGraphTokens(decrypt(session.graphRefreshTokenEnc));
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new GraphError(`Token refresh failed — re-authentication required: ${msg}`, 401);
  }

  const updated: Session = {
    ...session,
    graphAccessTokenEnc: encrypt(tokens.accessToken),
    graphRefreshTokenEnc: encrypt(tokens.refreshToken),
    graphExpiresAt: now + tokens.expiresIn * 1000,
  };
  await putSession(updated);

  if (VERBOSE) console.log(`[graph] refreshed token for session ${sessionId.slice(0, 8)}…`);
  return tokens.accessToken;
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
 * Make a Graph API request. Accepts the caller's opaque session id and
 * resolves it to a live Graph token (refreshing if needed) before calling.
 * Returns parsed JSON on success. Throws GraphError with HTTP status on failure.
 */
export async function graph<T = unknown>(
  sessionId: string,
  path: string,
  init: RequestInit = {}
): Promise<T> {
  const graphToken = await getGraphTokenForSession(sessionId);
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
 * Download binary content from Graph. Used by the download_file tool to read
 * an existing file's bytes for round-trip editing.
 *
 * Returns the raw bytes plus the server-reported content-type so the caller
 * can attach the correct MIME type to the MCP resource block. The Graph
 * /content endpoints normally 302 to a pre-authenticated download URL on
 * SharePoint storage; we follow that redirect automatically.
 */
export async function graphGetBinary(
  sessionId: string,
  path: string
): Promise<{ bytes: Uint8Array; contentType: string }> {
  const graphToken = await getGraphTokenForSession(sessionId);
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${graphToken}` },
    redirect: "follow",
  });

  if (VERBOSE) {
    console.log(`[graph] GET ${path} (binary) → ${res.status}`);
  }

  if (!res.ok) {
    let errBody: GraphErrorBody | undefined;
    try {
      errBody = (await res.json()) as GraphErrorBody;
    } catch {
      // not JSON
    }
    throw new GraphError(
      errBody?.error?.message ?? `Graph GET (binary) returned ${res.status}`,
      res.status,
      errBody?.error?.code,
      errBody?.error?.innerError?.["request-id"]
    );
  }

  const buffer = await res.arrayBuffer();
  return {
    bytes: new Uint8Array(buffer),
    contentType: res.headers.get("content-type") ?? "application/octet-stream",
  };
}

/**
 * Upload binary content to Graph. Used for file uploads under 4 MB.
 * Takes the caller's session id; resolves it to a Graph token internally.
 */
export async function graphPutBinary(
  sessionId: string,
  path: string,
  body: ArrayBuffer | Uint8Array | Blob,
  contentType: string = "application/octet-stream"
): Promise<unknown> {
  const graphToken = await getGraphTokenForSession(sessionId);
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
 * Takes the caller's session id; resolves it to a Graph token for the
 * session-create call. The actual chunked PUTs use the temporary uploadUrl
 * Microsoft returns, which is pre-authenticated so no token needed on those.
 *
 * https://learn.microsoft.com/graph/api/driveitem-createuploadsession
 */
export async function uploadFileViaSession(
  sessionId: string,
  uploadSessionPath: string,
  body: ArrayBuffer,
  conflictBehavior: "rename" | "replace" | "fail" = "replace"
): Promise<unknown> {
  // 1. Create the upload session (resolves the session id to a Graph token)
  const session = (await graph<{ uploadUrl: string }>(
    sessionId,
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
