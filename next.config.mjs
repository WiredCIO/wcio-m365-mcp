/** @type {import('next').NextConfig} */
const nextConfig = {
  // No UI build needed beyond the landing page
  output: undefined,
  experimental: {
    // mcp-handler uses long-lived streaming connections
    serverComponentsExternalPackages: ["@modelcontextprotocol/sdk"],
  },
};

export default nextConfig;
