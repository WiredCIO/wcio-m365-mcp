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
  WCIO_MCP_CLIENT_SECRET: z.string().min(20, "Client secret must be set for OBO flow"),
  WCIO_MCP_BASE_URL: z.string().url(),
  WCIO_MCP_VERBOSE_LOGGING: z.enum(["0", "1"]).default("0"),
});

function parseEnv() {
  const result = envSchema.safeParse({
    WCIO_MCP_TENANT_ID: process.env.WCIO_MCP_TENANT_ID,
    WCIO_MCP_CLIENT_ID: process.env.WCIO_MCP_CLIENT_ID,
    WCIO_MCP_CLIENT_SECRET: process.env.WCIO_MCP_CLIENT_SECRET,
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

/** Microsoft Entra v2 token endpoint for OBO exchanges. */
export const TOKEN_ENDPOINT = `https://login.microsoftonline.com/${env.WCIO_MCP_TENANT_ID}/oauth2/v2.0/token`;

/** Microsoft Graph API base. */
export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Expected audience for tokens issued to our MCP server.
 * Entra will use either the bare client ID or the Application ID URI
 * (api://<client-id>) depending on how the request was made.
 */
export const APP_AUDIENCE = [
  env.WCIO_MCP_CLIENT_ID,
  `api://${env.WCIO_MCP_CLIENT_ID}`,
];

/** Application ID URI — what we publish as `resource` in OAuth metadata. */
export const APP_ID_URI = `api://${env.WCIO_MCP_CLIENT_ID}`;

/** Our custom scope, exposed via "Expose an API" in Entra. */
export const APP_SCOPE = `${APP_ID_URI}/mcp.access`;

/** Verbose logging flag. */
export const VERBOSE = env.WCIO_MCP_VERBOSE_LOGGING === "1";
