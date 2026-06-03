/**
 * Environment variable parsing and validation.
 * Fails fast at module load if anything is missing or malformed.
 */
import { z } from "zod";

const uuidSchema = z.string().regex(
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
  "Must be a UUID"
);

const envSchema = z.object({
  WCIO_MCP_TENANT_ID: uuidSchema,
  WCIO_MCP_CLIENT_ID: uuidSchema,
  // Confidential client: required for the server-side authorization-code and
  // refresh-token exchanges with Entra.
  WCIO_MCP_CLIENT_SECRET: z.string().min(20, "Client secret must be set"),
  WCIO_MCP_BASE_URL: z.string().url(),
  // 32-byte key (base64-encoded) used to encrypt Entra refresh tokens at rest
  // in the KV store. Generate with: openssl rand -base64 32
  WCIO_MCP_TOKEN_ENCRYPTION_KEY: z
    .string()
    .refine((v) => Buffer.from(v, "base64").length === 32, {
      message: "Must be a base64-encoded 32-byte key (openssl rand -base64 32)",
    }),
  WCIO_MCP_VERBOSE_LOGGING: z.enum(["0", "1"]).default("0"),
});

// `next build` imports every route module to collect page data, which runs
// this file's top-level code. Build machines legitimately lack runtime secrets
// (client secret, encryption key), so we must not hard-fail there. The same
// validation re-runs on every serverless cold start — where NEXT_PHASE is unset
// and the real env IS present — so fail-fast behavior is preserved at runtime.
const IS_BUILD_PHASE = process.env.NEXT_PHASE === "phase-production-build";

type Env = z.infer<typeof envSchema>;

function parseEnv(): Env {
  const raw = {
    WCIO_MCP_TENANT_ID: process.env.WCIO_MCP_TENANT_ID,
    WCIO_MCP_CLIENT_ID: process.env.WCIO_MCP_CLIENT_ID,
    WCIO_MCP_CLIENT_SECRET: process.env.WCIO_MCP_CLIENT_SECRET,
    WCIO_MCP_BASE_URL: process.env.WCIO_MCP_BASE_URL,
    WCIO_MCP_TOKEN_ENCRYPTION_KEY: process.env.WCIO_MCP_TOKEN_ENCRYPTION_KEY,
    WCIO_MCP_VERBOSE_LOGGING: process.env.WCIO_MCP_VERBOSE_LOGGING,
  };

  const result = envSchema.safeParse(raw);

  if (!result.success) {
    if (IS_BUILD_PHASE) {
      // Defer the hard failure to runtime; supply harmless placeholders so the
      // derived constants below can be computed without throwing during build.
      return {
        WCIO_MCP_TENANT_ID:
          raw.WCIO_MCP_TENANT_ID ?? "00000000-0000-0000-0000-000000000000",
        WCIO_MCP_CLIENT_ID:
          raw.WCIO_MCP_CLIENT_ID ?? "00000000-0000-0000-0000-000000000000",
        WCIO_MCP_CLIENT_SECRET: raw.WCIO_MCP_CLIENT_SECRET ?? "",
        WCIO_MCP_BASE_URL: raw.WCIO_MCP_BASE_URL ?? "https://example.invalid",
        WCIO_MCP_TOKEN_ENCRYPTION_KEY:
          raw.WCIO_MCP_TOKEN_ENCRYPTION_KEY ?? Buffer.alloc(32).toString("base64"),
        WCIO_MCP_VERBOSE_LOGGING:
          raw.WCIO_MCP_VERBOSE_LOGGING === "1" ? "1" : "0",
      };
    }
    const issues = result.error.issues
      .map((i) => `  ${i.path.join(".")}: ${i.message}`)
      .join("\n");
    throw new Error(`Invalid environment configuration:\n${issues}`);
  }

  return result.data;
}

export const env = parseEnv();

/** Normalized base URL with no trailing slash. */
export const BASE_URL = env.WCIO_MCP_BASE_URL.replace(/\/+$/, "");

/** Microsoft Entra v2 OAuth issuer URL for our tenant. */
export const ISSUER = `https://login.microsoftonline.com/${env.WCIO_MCP_TENANT_ID}/v2.0`;

/** Microsoft Entra v2 authorization endpoint. */
export const ENTRA_AUTHORIZE_ENDPOINT = `https://login.microsoftonline.com/${env.WCIO_MCP_TENANT_ID}/oauth2/v2.0/authorize`;

/** Microsoft Entra v2 token endpoint (auth-code + refresh exchanges). */
export const ENTRA_TOKEN_ENDPOINT = `https://login.microsoftonline.com/${env.WCIO_MCP_TENANT_ID}/oauth2/v2.0/token`;

/** Microsoft Graph API base. */
export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Scopes we request from Entra on the server-side authorization-code leg.
 *
 * `.default` yields a Graph token containing every delegated permission the
 * app registration has admin-consented for the user, so we don't need to
 * enumerate individual Graph scopes here. The OIDC scopes give us an id_token
 * (for the user's identity) and a refresh token.
 *
 * IMPORTANT: we deliberately do NOT send an OAuth `resource` parameter — the
 * scope already encodes the Graph audience, and Entra v2.0 rejects requests
 * that carry both `resource` and a v2 scope (AADSTS9010010).
 */
export const GRAPH_SCOPES =
  "openid profile offline_access https://graph.microsoft.com/.default";

/** Our own callback URL, registered as the redirect URI in Entra. */
export const ENTRA_REDIRECT_URI = `${BASE_URL}/api/oauth/callback`;

/** Canonical resource identifier for this MCP server (RFC 8707 / RFC 9728). */
export const MCP_RESOURCE = `${BASE_URL}/api/mcp`;

/** Verbose logging flag. */
export const VERBOSE = env.WCIO_MCP_VERBOSE_LOGGING === "1";
