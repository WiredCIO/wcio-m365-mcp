/** @type {import('next').NextConfig} */
const nextConfig = {
  // No UI build needed beyond the landing page
  output: undefined,
  // mcp-handler uses long-lived streaming connections
  serverExternalPackages: ["@modelcontextprotocol/sdk"],
};

export default nextConfig;
