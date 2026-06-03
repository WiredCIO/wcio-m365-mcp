/**
 * Liveness check. Returns 200 with build info. Does not validate auth.
 * Useful for Vercel deployment health checks and uptime monitoring.
 */
import { NextResponse } from "next/server";
import { env } from "@/lib/env";

export const dynamic = "force-dynamic";

export async function GET() {
  return NextResponse.json({
    status: "ok",
    name: "wcio-m365-mcp",
    version: "0.1.0",
    tenantConfigured: Boolean(env.WCIO_MCP_TENANT_ID),
    baseUrl: env.WCIO_MCP_BASE_URL,
    timestamp: new Date().toISOString(),
  });
}
