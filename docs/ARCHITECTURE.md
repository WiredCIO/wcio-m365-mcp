# Architecture

This document describes how the Wired CIO M365 MCP server is built and why.
Read this before making structural changes.

---

## Goals

1. **Per-user identity**: every Graph API call runs as the user who initiated
   the AI conversation. Their name appears in audit logs. They can only
   access what they're permitted to access.
2. **No token storage on our server**: we are a stateless proxy. Tokens flow
   through; we never persist them. If our server is compromised, no historical
   credentials leak.
3. **Single deployment, org-wide use**: one Entra app, one Vercel deployment,
   one URL that anyone in the tenant can connect to.
4. **Composable with the existing read-only M365 connector**: the read side
   (search, read_resource) stays on Anthropic's published connector. We add
   the write side. Both can be active simultaneously in Claude.

---

## Request flow

### Initial OAuth setup (once per user, per device)

```
User           Claude.ai             Our MCP            Microsoft Entra
 │                │                    │                    │
 │ "add connector"│                    │                    │
 │───────────────►│                    │                    │
 │                │ GET /.well-known/oauth-protected-resource
 │                │───────────────────►│                    │
 │                │  authorization_servers: [Entra/{tid}/v2.0]
 │                │◄───────────────────│                    │
 │                │                    │                    │
 │                │ GET /.well-known/openid-configuration   │
 │                │────────────────────────────────────────►│
 │                │  authorization_endpoint, token_endpoint │
 │                │◄────────────────────────────────────────│
 │                │                    │                    │
 │   browser redirect to authorization_endpoint w/ PKCE     │
 │◄─────────────────────────────────────────────────────────│
 │   sign in (Microsoft Entra)                              │
 │─────────────────────────────────────────────────────────►│
 │                │                    │                    │
 │   redirect to claude.ai/api/mcp/auth_callback?code=...   │
 │◄─────────────────────────────────────────────────────────│
 │                │ POST token_endpoint w/ code + PKCE      │
 │                │────────────────────────────────────────►│
 │                │  access_token, refresh_token            │
 │                │◄────────────────────────────────────────│
 │                │                    │                    │
 │                │ stores tokens in Claude.ai's user storage
 │                │                    │                    │
```

### Tool invocation (every MCP request)

```
Claude.ai                  Our MCP                Microsoft Graph
    │                        │                        │
    │ POST /api/mcp/mcp       │                        │
    │ Authorization: Bearer X │                        │
    │────────────────────────►│                        │
    │                        │                        │
    │     verifyEntraToken(X):                         │
    │      - fetch JWKS from Entra (cached)            │
    │      - verify signature                          │
    │      - check iss = our tenant                    │
    │      - check aud = Graph                         │
    │      - check tid = our tenant                    │
    │     → AuthInfo {token: X, scopes, userId}        │
    │                        │                        │
    │     tool handler called with AuthInfo            │
    │     extracts accessToken                         │
    │                        │                        │
    │                        │ Graph API call         │
    │                        │ Authorization: Bearer X │
    │                        │───────────────────────►│
    │                        │  result                │
    │                        │◄───────────────────────│
    │                        │                        │
    │  MCP tool result        │                        │
    │◄────────────────────────│                        │
```

When the token expires (typically 1 hour), Claude.ai uses the refresh_token
to get a new one transparently. The MCP server doesn't participate in refresh
— it just validates whatever token shows up in the next call.

---

## Code layout

```
app/
├── api/
│   ├── mcp/[transport]/route.ts       # MCP HTTP endpoint (POST/GET/DELETE)
│   └── health/route.ts                # liveness check
├── .well-known/
│   └── oauth-protected-resource/      # OAuth metadata for client discovery
│       └── route.ts
├── layout.tsx                         # Next.js shell
└── page.tsx                           # public landing page

lib/
├── env.ts                             # env parsing, constants
├── auth.ts                            # JWT verification, AuthInfo
├── graph.ts                           # fetch helpers for Graph API
└── tools/
    ├── index.ts                       # tool registration aggregator
    ├── files.ts                       # OneDrive / SharePoint write tools
    ├── mail.ts                        # Outlook send/draft/reply
    ├── calendar.ts                    # event create/update/cancel
    └── teams.ts                       # chat/channel posts
```

### Why Next.js?

Vercel's `mcp-handler` package is designed around Next.js App Router route
handlers. We get:
- Automatic streaming HTTP support (both SSE and Streamable HTTP transports)
- Per-route runtime config (maxDuration, dynamic)
- Simple static file serving for `/.well-known/*`
- Familiar deployment model for anyone who knows Next on Vercel

We don't use the React rendering features other than the minimal landing page.

### Why no client secret?

OAuth 2.1 with PKCE doesn't require a confidential client. PKCE protects the
authorization code from interception. A confidential client would buy us
client authentication during the token exchange — but since our server isn't
exchanging codes (Claude.ai is), the protection wouldn't apply to us anyway.

This also means there's no secret to leak from our codebase. The Entra
client ID is public information.

### Why pass-through tokens instead of on-behalf-of (OBO)?

OBO would let our server act with an extended audience by exchanging the
user's token for a different one. We'd need a client secret for that, plus
additional Entra config. For our use case (Graph API only), pass-through is
strictly simpler and equally functional. If we later needed to call a
different API beyond Graph, OBO might become attractive.

---

## Security model

### What the server can do
- Validate tokens
- Forward authenticated requests to Graph
- Log non-sensitive metadata (user identifier, tool name, success/failure)

### What the server cannot do
- Act on behalf of any user without their token in the current request
- Generate or refresh tokens
- Access OneDrive/SharePoint/Outlook outside of a user's Graph-scoped permissions
- Store user data persistently — every container restart is clean state

### What attackers can attempt
- **Forged tokens**: blocked by JWT signature verification against Microsoft's JWKS.
- **Token replay across tenants**: blocked by `tid` claim check.
- **Token misuse for wrong API**: blocked by `aud` claim check (must be Graph).
- **Compromising our server to grab tokens**: limited blast radius — we only see tokens during their ~1-hour validity, no refresh tokens, no client secret. Tokens are not logged.
- **Prompt injection from a file/email content**: this is the AI client's responsibility, not ours. The client is expected to confirm with the user before destructive actions.

### Audit story

- **Who did what**: Microsoft Entra sign-in logs + Graph activity logs, indexed by user UPN.
- **What our server did**: optional verbose logging captures tool calls without tokens. Production default is off.
- **Code changes**: GitHub commit history under the wiredcio org.
- **Deployment changes**: Vercel deployment history.

---

## Adding new tools

1. Decide which Graph permissions the new tool needs. Add them to
   [ENTRA_SETUP.md](./ENTRA_SETUP.md) and run **Grant admin consent** in Entra.
2. Add a registration function in `lib/tools/<category>.ts` (or create a new
   file if it's a new category).
3. Use `graph()` or `graphPutBinary()` for the Graph call. Don't reinvent the
   error handling — the helpers do it consistently.
4. Define inputs with Zod for runtime validation + auto-generated tool schemas.
5. Register in `lib/tools/index.ts`.
6. Test locally with `npm run dev`, then push and verify via the Vercel
   preview deployment.

### Tool design principles

- **Reversible actions are safer than destructive ones**: prefer
  `create_draft` over `send_email` when the use case allows; expose both so
  the user can choose.
- **Identify items by path AND ID**: paths are user-friendly, IDs are stable.
- **Use OneDrive's versioning**: any tool that modifies file content should
  rely on OneDrive's automatic version retention rather than trying to
  manage versions ourselves.
- **Surface Graph errors literally**: don't try to re-interpret a 403. Let
  the user see the Graph error message and they can act on it.

---

## Future work

Things explicitly not in v1 that we may want later:

- **Approval gates**: a tool that posts a draft to a review channel and waits
  for a human to approve before executing send_email. Requires state.
- **Bulk operations**: batched file uploads with progress reporting.
- **On-Behalf-Of flow**: if we add a second API beyond Graph.
- **App-only tokens** for system tasks like scheduled cleanup. Higher risk
  surface; not needed for any current use case.
- **Webhook support**: subscribe to Graph notifications and trigger tools.
  Would require persistent state.
