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
 * Make a Graph API request. Returns parsed JSON on success.
 * Throws GraphError with HTTP status and Graph error code on failure.
 */
export async function graph<T = unknown>(
  accessToken: string,
  path: string,
  init: RequestInit = {}
): Promise<T> {
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${accessToken}`);
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

  // Some Graph endpoints return 204 No Content
  if (response.status === 204) {
    return undefined as T;
  }

  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return (await response.json()) as T;
  }
  // Caller should handle non-JSON via raw fetch
  return undefined as T;
}

/**
 * Upload binary content to Graph. Used for file uploads under 4 MB.
 * For larger files, see uploadFileViaSession.
 */
export async function graphPutBinary(
  accessToken: string,
  path: string,
  body: ArrayBuffer | Uint8Array | Blob,
  contentType: string = "application/octet-stream"
): Promise<unknown> {
  return graph(accessToken, path, {
    method: "PUT",
    headers: { "Content-Type": contentType },
    body: body as BodyInit,
  });
}

/**
 * Upload a file >4 MB using an upload session.
 *
 * https://learn.microsoft.com/graph/api/driveitem-createuploadsession
 *
 * Strategy:
 *  1. Create upload session — Graph returns an uploadUrl
 *  2. PUT chunks of up to 60 MB to that URL with Content-Range headers
 *  3. Final chunk returns the DriveItem metadata
 */
export async function uploadFileViaSession(
  accessToken: string,
  uploadSessionPath: string,
  body: ArrayBuffer,
  conflictBehavior: "rename" | "replace" | "fail" = "replace"
): Promise<unknown> {
  // 1. Create the upload session
  const session = (await graph<{ uploadUrl: string }>(
    accessToken,
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
