import type { ReactNode } from "react";

export const metadata = {
  title: "Wired CIO M365 MCP",
  description: "MCP server providing write access to Microsoft 365 for AI assistants",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body
        style={{
          margin: 0,
          fontFamily:
            "-apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif",
          color: "#0C4651",
          backgroundColor: "#FAFAFA",
        }}
      >
        {children}
      </body>
    </html>
  );
}
