/**
 * OAuth 2.0 Protected Resource Metadata (RFC 9728).
 *
 * MCP clients fetch this to discover where to authenticate. We point them
 * at Microsoft's tenant-specific v2.0 authorization server, which has its
 * own metadata at .../.well-known/openid-configuration.
 *
 * Returned in response to:
 *   1. Direct GET /.well-known/oauth-protected-resource
 *   2. 401 responses from /api/mcp/* (via the WWW-Authenticate header that
 *      mcp-handler emits, which contains the resource_metadata URL)
 */
import {
  protectedResourceHandler,
  metadataCorsOptionsRequestHandler,
} from "mcp-handler";
import { ISSUER } from "@/lib/env";

const handler = protectedResourceHandler({
  authServerUrls: [ISSUER],
});

const corsHandler = metadataCorsOptionsRequestHandler();

export { handler as GET, corsHandler as OPTIONS };
