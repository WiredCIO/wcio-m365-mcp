/**
 * Main MCP endpoint.
 *
 * Exposes the tools at `/api/mcp/[transport]` where [transport] is either
 * `mcp` (Streamable HTTP) or `sse` (legacy Server-Sent Events). mcp-handler
 * dispatches based on that path segment.
 *
 * Every request is wrapped in withMcpAuth, which validates the Bearer token
 * via verifyEntraToken. Failed auth returns 401 with a WWW-Authenticate
 * header pointing at our protected-resource metadata endpoint, which in turn
 * points MCP clients (like Claude.ai) at Microsoft's authorization server.
 */
import { createMcpHandler, withMcpAuth } from "mcp-handler";
import { verifyEntraToken } from "@/lib/auth";
import { registerAllTools } from "@/lib/tools";

export const dynamic = "force-dynamic";
export const maxDuration = 800;

const handler = createMcpHandler(
  (server) => {
    registerAllTools(server);
  },
  {
    // Server identity advertised to MCP clients
    serverInfo: {
      name: "wcio-m365-mcp",
      version: "0.1.0",
    },
  },
  {
    basePath: "/api",
  }
);

const authHandler = withMcpAuth(handler, verifyEntraToken, {
  required: true,
  // We don't enforce specific scopes at the transport layer — each tool can
  // check the scopes it needs from the AuthInfo passed in its handler context.
  // This keeps tools self-documenting and allows per-tool scope requirements
  // (e.g. send_email needs Mail.Send, list_folder only needs Files.Read).
  resourceMetadataPath: "/.well-known/oauth-protected-resource",
});

export { authHandler as GET, authHandler as POST, authHandler as DELETE };
