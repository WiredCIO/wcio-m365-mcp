/**
 * OneDrive and SharePoint file operation tools.
 *
 * For OneDrive (user's personal drive): identify files by path or itemId
 * relative to /me/drive.
 * For SharePoint sites: identify files via a siteId parameter, then path
 * within that site's default document library.
 */
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import {
  graph,
  graphPutBinary,
  uploadFileViaSession,
  userDrivePath,
  userDriveItemById,
  siteDrivePath,
  GraphError,
} from "../graph";

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  parentReference?: { driveId?: string; path?: string };
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  lastModifiedDateTime?: string;
}

/** 4 MB threshold for switching from simple PUT to upload session. */
const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;

function resolveItemRef(opts: {
  siteId?: string;
  path?: string;
  itemId?: string;
}): string {
  if (opts.itemId) {
    return opts.siteId
      ? `/sites/${opts.siteId}/drive/items/${opts.itemId}`
      : userDriveItemById(opts.itemId);
  }
  if (opts.path) {
    return opts.siteId ? siteDrivePath(opts.siteId, opts.path) : userDrivePath(opts.path);
  }
  throw new Error("Either `path` or `itemId` is required");
}

function formatItem(item: DriveItem): string {
  const kind = item.folder ? "folder" : "file";
  const size = item.size ? ` (${formatBytes(item.size)})` : "";
  const url = item.webUrl ? `\n  URL: ${item.webUrl}` : "";
  return `${kind}: ${item.name}${size}\n  ID: ${item.id}${url}`;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

export function registerFileTools(server: McpServer) {
  // ==========================================================================
  // upload_file — create a new file
  // ==========================================================================
  server.tool(
    "upload_file",
    "Upload a new file to OneDrive or a SharePoint site. Content is provided " +
      "as base64-encoded bytes. For files over 4 MB, an upload session is used " +
      "automatically. If a file already exists at the path, the conflict_behavior " +
      "parameter controls whether to replace it (default), rename, or fail.",
    {
      path: z
        .string()
        .describe(
          "Destination path including filename, relative to drive root. " +
            "Example: 'Documents/Reports/Q3-summary.xlsx'"
        ),
      content_base64: z.string().describe("File content as base64 string"),
      content_type: z
        .string()
        .default("application/octet-stream")
        .describe("MIME type. Common: application/pdf, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, image/png"),
      site_id: z
        .string()
        .optional()
        .describe(
          "SharePoint site ID. Omit to upload to the calling user's OneDrive. " +
            "Get site IDs via sharepoint_search on the existing read-only connector."
        ),
      conflict_behavior: z
        .enum(["rename", "replace", "fail"])
        .default("replace")
        .describe("What to do if a file already exists at that path"),
    },
    async ({ path, content_base64, content_type, site_id, conflict_behavior }, { authInfo }) => {
      const token = authInfo!.extra!.accessToken as string;
      const bytes = Buffer.from(content_base64, "base64");
      const itemPath = resolveItemRef({ siteId: site_id, path });

      try {
        let result: DriveItem;
        if (bytes.byteLength <= SIMPLE_UPLOAD_LIMIT) {
          // Simple PUT to /content with a conflict header
          const contentPath = `${itemPath}:/content?@microsoft.graph.conflictBehavior=${conflict_behavior}`;
          result = (await graphPutBinary(token, contentPath, bytes, content_type)) as DriveItem;
        } else {
          // Upload session for big files
          const sessionPath = `${itemPath}:/createUploadSession`;
          result = (await uploadFileViaSession(
            token,
            sessionPath,
            bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
            conflict_behavior
          )) as DriveItem;
        }
        return {
          content: [
            {
              type: "text",
              text: `Uploaded successfully.\n${formatItem(result)}`,
            },
          ],
        };
      } catch (err) {
        return errorResult("upload_file", err);
      }
    }
  );

  // ==========================================================================
  // update_file — replace existing file content
  // ==========================================================================
  server.tool(
    "update_file",
    "Replace the content of an existing file. OneDrive automatically retains the " +
      "previous version in version history, so the prior content is recoverable. " +
      "Identify the target by path OR by itemId.",
    {
      path: z
        .string()
        .optional()
        .describe("Path to existing file. Either path OR item_id required."),
      item_id: z
        .string()
        .optional()
        .describe("Graph item ID of the existing file. Either path OR item_id required."),
      content_base64: z.string().describe("New file content as base64 string"),
      content_type: z
        .string()
        .default("application/octet-stream")
        .describe("MIME type of the new content"),
      site_id: z.string().optional().describe("SharePoint site ID, if not OneDrive"),
    },
    async ({ path, item_id, content_base64, content_type, site_id }, { authInfo }) => {
      if (!path && !item_id) {
        return {
          content: [{ type: "text", text: "Error: must provide either `path` or `item_id`" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;
      const bytes = Buffer.from(content_base64, "base64");
      const itemRef = resolveItemRef({ siteId: site_id, path, itemId: item_id });

      try {
        let result: DriveItem;
        if (bytes.byteLength <= SIMPLE_UPLOAD_LIMIT) {
          // For an existing item identified by ID, the content path is /content
          // For path-based, we use :/content
          const contentPath = item_id ? `${itemRef}/content` : `${itemRef}:/content`;
          result = (await graphPutBinary(token, contentPath, bytes, content_type)) as DriveItem;
        } else {
          const sessionPath = item_id
            ? `${itemRef}/createUploadSession`
            : `${itemRef}:/createUploadSession`;
          result = (await uploadFileViaSession(
            token,
            sessionPath,
            bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
            "replace"
          )) as DriveItem;
        }
        return {
          content: [
            {
              type: "text",
              text: `Replaced file content. Prior version retained in OneDrive history.\n${formatItem(result)}`,
            },
          ],
        };
      } catch (err) {
        return errorResult("update_file", err);
      }
    }
  );

  // ==========================================================================
  // list_folder
  // ==========================================================================
  server.tool(
    "list_folder",
    "List the contents of a folder in OneDrive or a SharePoint site.",
    {
      path: z
        .string()
        .default("")
        .describe("Folder path. Empty string lists the drive root."),
      item_id: z.string().optional().describe("Alternative: folder item ID"),
      site_id: z.string().optional().describe("SharePoint site ID, if not OneDrive"),
      top: z.number().int().min(1).max(200).default(50).describe("Max items to return"),
    },
    async ({ path, item_id, site_id, top }, { authInfo }) => {
      const token = authInfo!.extra!.accessToken as string;
      let listPath: string;
      if (item_id) {
        listPath = site_id
          ? `/sites/${site_id}/drive/items/${item_id}/children`
          : `/me/drive/items/${item_id}/children`;
      } else if (!path) {
        listPath = site_id ? `/sites/${site_id}/drive/root/children` : `/me/drive/root/children`;
      } else {
        const base = site_id ? siteDrivePath(site_id, path) : userDrivePath(path);
        listPath = `${base}:/children`;
      }
      listPath += `?$top=${top}`;

      try {
        const result = (await graph(token, listPath)) as { value: DriveItem[] };
        if (!result.value.length) {
          return { content: [{ type: "text", text: "Folder is empty." }] };
        }
        const text = result.value.map(formatItem).join("\n\n");
        return { content: [{ type: "text", text }] };
      } catch (err) {
        return errorResult("list_folder", err);
      }
    }
  );

  // ==========================================================================
  // create_folder
  // ==========================================================================
  server.tool(
    "create_folder",
    "Create a new folder. Parent must exist. Set conflict_behavior=rename to " +
      "auto-rename if a folder with the same name already exists.",
    {
      parent_path: z
        .string()
        .default("")
        .describe("Path of parent folder. Empty for drive root."),
      folder_name: z.string().describe("Name of the new folder"),
      site_id: z.string().optional().describe("SharePoint site ID, if not OneDrive"),
      conflict_behavior: z.enum(["rename", "replace", "fail"]).default("fail"),
    },
    async ({ parent_path, folder_name, site_id, conflict_behavior }, { authInfo }) => {
      const token = authInfo!.extra!.accessToken as string;
      let parentPath: string;
      if (!parent_path) {
        parentPath = site_id ? `/sites/${site_id}/drive/root/children` : `/me/drive/root/children`;
      } else {
        const base = site_id ? siteDrivePath(site_id, parent_path) : userDrivePath(parent_path);
        parentPath = `${base}:/children`;
      }

      try {
        const result = (await graph(token, parentPath, {
          method: "POST",
          body: JSON.stringify({
            name: folder_name,
            folder: {},
            "@microsoft.graph.conflictBehavior": conflict_behavior,
          }),
        })) as DriveItem;
        return {
          content: [{ type: "text", text: `Created folder.\n${formatItem(result)}` }],
        };
      } catch (err) {
        return errorResult("create_folder", err);
      }
    }
  );

  // ==========================================================================
  // delete_item
  // ==========================================================================
  server.tool(
    "delete_item",
    "Move an item (file or folder) to the recycle bin. Recoverable from OneDrive " +
      "or SharePoint trash for 30 days. Identify by path or item_id.",
    {
      path: z.string().optional(),
      item_id: z.string().optional(),
      site_id: z.string().optional(),
    },
    async ({ path, item_id, site_id }, { authInfo }) => {
      if (!path && !item_id) {
        return {
          content: [{ type: "text", text: "Error: provide either `path` or `item_id`" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;
      const itemRef = resolveItemRef({ siteId: site_id, path, itemId: item_id });

      try {
        await graph(token, itemRef, { method: "DELETE" });
        return { content: [{ type: "text", text: `Moved to recycle bin: ${path ?? item_id}` }] };
      } catch (err) {
        return errorResult("delete_item", err);
      }
    }
  );

  // ==========================================================================
  // get_item_metadata
  // ==========================================================================
  server.tool(
    "get_item_metadata",
    "Get metadata for a file or folder including ID, size, last-modified date, " +
      "version history pointer, and webUrl. Useful before update_file to confirm you " +
      "have the right target.",
    {
      path: z.string().optional(),
      item_id: z.string().optional(),
      site_id: z.string().optional(),
    },
    async ({ path, item_id, site_id }, { authInfo }) => {
      if (!path && !item_id) {
        return {
          content: [{ type: "text", text: "Error: provide either `path` or `item_id`" }],
          isError: true,
        };
      }
      const token = authInfo!.extra!.accessToken as string;
      const itemRef = resolveItemRef({ siteId: site_id, path, itemId: item_id });
      try {
        const result = (await graph(token, itemRef)) as DriveItem;
        return {
          content: [
            {
              type: "text",
              text:
                formatItem(result) +
                (result.lastModifiedDateTime
                  ? `\n  Last modified: ${result.lastModifiedDateTime}`
                  : ""),
            },
          ],
        };
      } catch (err) {
        return errorResult("get_item_metadata", err);
      }
    }
  );

  // ==========================================================================
  // move_item
  // ==========================================================================
  server.tool(
    "move_item",
    "Move a file or folder to a different parent folder. Optionally rename in flight.",
    {
      item_id: z.string().describe("ID of the item to move"),
      destination_path: z
        .string()
        .describe("New parent folder path (the folder, not including the item name)"),
      new_name: z.string().optional().describe("Optional new name for the item"),
      site_id: z.string().optional(),
    },
    async ({ item_id, destination_path, new_name, site_id }, { authInfo }) => {
      const token = authInfo!.extra!.accessToken as string;
      const itemRef = resolveItemRef({ siteId: site_id, itemId: item_id });
      const parentRef = site_id
        ? siteDrivePath(site_id, destination_path)
        : userDrivePath(destination_path);

      try {
        const body: Record<string, unknown> = {
          parentReference: { path: parentRef.replace(":/", "/root:/") },
        };
        if (new_name) body.name = new_name;

        const result = (await graph(token, itemRef, {
          method: "PATCH",
          body: JSON.stringify(body),
        })) as DriveItem;
        return { content: [{ type: "text", text: `Moved.\n${formatItem(result)}` }] };
      } catch (err) {
        return errorResult("move_item", err);
      }
    }
  );
}

function errorResult(toolName: string, err: unknown) {
  if (err instanceof GraphError) {
    return {
      content: [
        {
          type: "text" as const,
          text:
            `Microsoft Graph error in ${toolName}: ${err.message}\n` +
            `HTTP ${err.statusCode}${err.graphCode ? ` (${err.graphCode})` : ""}` +
            (err.requestId ? `\nRequest ID: ${err.requestId}` : ""),
        },
      ],
      isError: true,
    };
  }
  const msg = err instanceof Error ? err.message : String(err);
  return {
    content: [{ type: "text" as const, text: `Error in ${toolName}: ${msg}` }],
    isError: true,
  };
}
