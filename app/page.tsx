/**
 * Public landing page. Anyone visiting the root URL gets a brief explanation
 * of what this service is. No sensitive info is rendered.
 */
import { env } from "@/lib/env";

export default function Home() {
  return (
    <main
      style={{
        maxWidth: 720,
        margin: "0 auto",
        padding: "4rem 1.5rem",
        lineHeight: 1.6,
      }}
    >
      <header style={{ borderBottom: "2px solid #2CC295", paddingBottom: "1rem", marginBottom: "2rem" }}>
        <h1 style={{ margin: 0, color: "#0C4651" }}>Wired CIO M365 MCP</h1>
        <p style={{ margin: "0.5rem 0 0 0", color: "#666" }}>
          Microsoft 365 write capability for AI assistants via Model Context Protocol
        </p>
      </header>

      <section>
        <h2 style={{ color: "#0C4651" }}>What this is</h2>
        <p>
          This service exposes Microsoft 365 actions (OneDrive, SharePoint, Outlook,
          Calendar, Teams) as Model Context Protocol tools. AI assistants that connect
          to this server can read and write Microsoft 365 data on behalf of the
          authenticated user.
        </p>
        <p>
          Every action runs under the calling user's own Microsoft identity. The
          server is stateless and does not store tokens.
        </p>
      </section>

      <section>
        <h2 style={{ color: "#0C4651" }}>For Wired CIO users</h2>
        <p>To connect from Claude.ai:</p>
        <ol>
          <li>Settings → Integrations → Add custom connector</li>
          <li>
            Server URL:{" "}
            <code style={{ background: "#F4F4F4", padding: "2px 6px", borderRadius: 4 }}>
              {env.WCIO_MCP_BASE_URL}/api/mcp/mcp
            </code>
          </li>
          <li>Sign in with your @wiredcio.com Microsoft account when prompted</li>
        </ol>
      </section>

      <footer style={{ marginTop: "3rem", paddingTop: "1rem", borderTop: "1px solid #E0E0E0", fontSize: "0.85rem", color: "#888" }}>
        <p>
          Wired CIO · 2921 N Milwaukee Ave Suite 1, Chicago, IL 60618 · (312) 210-0318 ·{" "}
          <a href="https://wiredcio.com" style={{ color: "#0C4651" }}>
            wiredcio.com
          </a>
        </p>
      </footer>
    </main>
  );
}
