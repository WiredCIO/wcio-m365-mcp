# Wired CIO Microsoft 365 MCP

[![Status: scaffolded](https://img.shields.io/badge/status-scaffolded-2CC295)](./docs/ROLLOUT.md)
[![Deploy: Vercel](https://img.shields.io/badge/deploy-Vercel-0C4651)](./docs/ROLLOUT.md)

Microsoft 365 write capability for AI assistants, exposed via Model Context Protocol.

The existing read-only M365 connector Anthropic publishes covers `sharepoint_search`, `outlook_email_search`, and `read_resource` — useful for retrieving data. This server adds the missing write side: uploading and replacing files in OneDrive and SharePoint, sending and drafting mail, creating calendar events, posting to Teams chats and channels.

Deployed on Vercel. Single-tenant to `wiredcio.com`. Each user authenticates with their own Microsoft account — every action shows up under their name in Entra and Graph audit logs.

---

## Architecture at a glance

```
┌──────────────┐         ┌─────────────────────────┐         ┌────────────────────┐
│  Claude.ai   │ MCP/OAuth│  This server (Vercel)  │ Graph   │  Microsoft Graph   │
│  (each user) ├─────────►│  Stateless, no storage  ├────────►│  OneDrive, SP,     │
│              │  Bearer  │  Token passthrough      │  Bearer │  Outlook, Teams    │
└──────────────┘  token   └─────────────────────────┘  token  └────────────────────┘
                                       ▲
                                       │ OAuth metadata discovery
                                       ▼
                          ┌─────────────────────────┐
                          │  Entra ID app (WCIO)    │
                          │  Admin consent granted  │
                          └─────────────────────────┘
```

The server validates each request's Bearer token against Microsoft's JWKS, checks `iss`/`aud`/`tid` claims, and forwards the same token to Graph. No tokens are stored, logged, or transformed.

Full details: [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md)

---

## Tool surface

**Files** (OneDrive + SharePoint)
- `upload_file` — create a new file (handles >4 MB via upload session)
- `update_file` — replace existing file content; OneDrive auto-versions
- `list_folder` — enumerate folder contents
- `create_folder`
- `delete_item` — soft delete to recycle bin
- `move_item`
- `get_item_metadata`

**Mail** (Outlook)
- `send_email` — send immediately
- `create_draft` — leave in Drafts for user review
- `reply_to_email` — threaded reply or reply-all

**Calendar**
- `create_event` — including Teams meeting auto-generation
- `update_event` — partial updates by field
- `cancel_event` — sends cancellation notice

**Teams**
- `post_chat_message` — 1:1 or group chat
- `post_channel_message` — to a team channel
- `reply_to_channel_message` — threaded reply

---

## Setup

Roughly 45 minutes end to end:

1. **Push the code** — create a private repo under the `wiredcio` GitHub org and push this folder.
2. **Register the Entra app** — follow [docs/ENTRA_SETUP.md](./docs/ENTRA_SETUP.md). Grants delegated Graph permissions and admin consent.
3. **Deploy to Vercel** — connect the repo, set three environment variables, deploy. See [docs/ROLLOUT.md](./docs/ROLLOUT.md).
4. **Add the connector in Claude.ai** — URL is `https://<your-deployment>/api/mcp/mcp`, sign in with your Microsoft account.
5. **Roll out to the team** — send the email template in ROLLOUT.md.

---

## Local development

```bash
npm install
cp .env.example .env.local
# fill in tenant ID, client ID, base URL
npm run dev
```

Server runs on `http://localhost:3000`. Health check at `/api/health`. OAuth metadata at `/.well-known/oauth-protected-resource`. MCP endpoint at `/api/mcp/mcp`.

For Claude.ai to authenticate against a local server, you'll need a tunnel (e.g. ngrok or `vercel dev` with a preview deployment). Easier to test against an actual preview deployment.

---

## Security model

- **No client secret** — uses OAuth 2.1 with PKCE. Nothing sensitive in the codebase or env vars beyond the public app ID and tenant ID.
- **No token storage** — tokens are validated and forwarded, never persisted.
- **Per-user identity** — Graph audit logs show real user names, not a service account.
- **Tenant pinning** — only tokens issued for `wiredcio.com` are accepted. Defense-in-depth checks both issuer and `tid` claim.
- **Stateless** — every container restart is clean state. No data exfil surface beyond the live request.

Threat model and audit story: [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md#security-model)

---

## Repository layout

```
wcio-m365-mcp/
├── README.md                   you are here
├── package.json
├── tsconfig.json
├── next.config.mjs
├── vercel.json
├── .env.example
├── .gitignore
├── app/                        Next.js App Router
│   ├── api/mcp/[transport]/route.ts     MCP endpoint (POST/GET/DELETE)
│   ├── api/health/route.ts              liveness check
│   ├── .well-known/oauth-protected-resource/route.ts
│   ├── layout.tsx                       minimal shell
│   └── page.tsx                         public landing page
├── lib/
│   ├── env.ts                           env validation + constants
│   ├── auth.ts                          JWT verification
│   ├── graph.ts                         Graph fetch helpers
│   └── tools/                           tool implementations
│       ├── index.ts
│       ├── files.ts
│       ├── mail.ts
│       ├── calendar.ts
│       └── teams.ts
└── docs/
    ├── ENTRA_SETUP.md                   step-by-step Entra config
    ├── ROLLOUT.md                       deployment + org rollout
    └── ARCHITECTURE.md                  for future maintainers
```

---

## Maintenance

- **Adding a new tool**: see [ARCHITECTURE.md → Adding new tools](./docs/ARCHITECTURE.md#adding-new-tools).
- **Adding a new Graph scope**: update [ENTRA_SETUP.md](./docs/ENTRA_SETUP.md) → API permissions, run admin consent again, then ship code that uses it.
- **Deployment changes**: every push to `main` auto-deploys. Use feature branches + Vercel preview URLs for sensitive changes.
- **Disaster recovery**: the server holds no state. Redeploy anywhere that runs Next.js. See [ROLLOUT.md → Disaster recovery](./docs/ROLLOUT.md#disaster-recovery).

---

## License

Internal Wired CIO project. Not for external distribution.

Contact: sales@wiredcio.com · (312) 210-0318
