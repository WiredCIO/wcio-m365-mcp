/**
 * Aggregates every tool module into a single registration function called
 * from the route handler. This keeps the route file small and lets us add
 * new tool groups without editing the wiring.
 */
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerFileTools } from "./files";
import { registerMailTools } from "./mail";
import { registerCalendarTools } from "./calendar";
import { registerTeamsTools } from "./teams";

export function registerAllTools(server: McpServer) {
  registerFileTools(server);
  registerMailTools(server);
  registerCalendarTools(server);
  registerTeamsTools(server);
}
