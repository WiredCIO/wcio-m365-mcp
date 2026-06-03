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
  WCIO_MCP_BASE_URL: z.string().url(),
  WCIO_MCP_VERBOSE_LOGGING: z.enum(["0", "1"]).default("0"),
});

function parseEnv() {
  const result = envSchema.safeParse({
    WCIO_MCP_TENANT_ID: process.env.WCIO_MCP_TENANT_ID,
    WCIO_MCP_CLIENT_ID: process.env.WCIO_MCP_CLIENT_ID,
    WCIO_MCP_BASE_URL: process.env.WCIO_MCP_BASE_URL,
    WCIO_MCP_VERBOSE_LOGGING: process.env.WCIO_MCP_VERBOSE_LOGGING,
  });

  if (!result.success) {
    const issues = result.error.issues
      .map((i) => `  ${i.path.join(".")}: ${i.message}`)
      .join("\n");
    throw new Error(`Invalid environment configuration:\n${issues}`);
  }

  return result.data;
}

export const env = parseEnv();

/** Microsoft Entra v2 OAuth issuer URL for our tenant. */
export const ISSUER = `https://login.microsoftonline.com/${env.WCIO_MCP_TENANT_ID}/v2.0`;

/** Microsoft Graph API base. */
export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** Expected audience for Graph-scoped tokens. */
export const GRAPH_AUDIENCE = [
  "https://graph.microsoft.com",
  "00000003-0000-0000-c000-000000000000",
];

/** Verbose logging flag. */
export const VERBOSE = env.WCIO_MCP_VERBOSE_LOGGING === "1";
