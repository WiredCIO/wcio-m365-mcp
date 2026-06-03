# Architecture

This document describes how the Wired CIO M365 MCP server is built and why.
Read this before making structural changes.

---

## Goals

1. **Per-user identity**: every Graph API call runs as the user who initiated
   the AI conversation. Their name appears in audit logs. They can only
   access what they're permitted to access.
2. **Be our own OAuth authorization server**: Claude talks OAuth only to us;
   we talk to Microsoft Entra server-side. This is what lets us work around
   two Entra limitations at once — no Dynamic Client Registration, and the
   `resource`+`scope` rejection (AADSTS9010010). Entra tokens never reach
   Claude; they live encrypted on our side and we hand Claude an opaque Bearer.
3. **Single deployment, org-wide use**: one Entra app, one Vercel deployment,
   one URL that anyone in the tenant can connect to.
4. **Composable with the existing read-only M365 connector**: the read side
   (search, read_resource) stays on Anthropic's published connector. We add
   the write side. Both can be active simultaneously in Claude.

---

## Request flow

We act as an OAuth 2.1 authorization server (Authorization Code + PKCE, with
Dynamic Client Registration) that proxies Microsoft Entra. Claude registers
with us, authorizes against us, and exchanges codes with us. Behind the
scenes we run the real authorization-code and refresh-token exchanges with
Entra as a confidential client, and we keep the Entra tokens encrypted in
Redis. Claude only ever holds an opaque Bearer that maps to a server session.

### Initial OAuth setup (once per user, per device)

```
User        Claude.ai            Our MCP (Auth Server)        Microsoft Entra
 │             │                       │                          │
 │"add conn."  │                       │                          │
 │────────────►│                       │                          │
 │             │ GET /.well-known/oauth-protected-resource        │
 │             │──────────────────────►│                          │
 │             │  authorization_servers: [<our base URL>]         │
 │             │◄──────────────────────│                          │
 │             │ GET /.well-known/oauth-authorization-server      │
 │             │──────────────────────►│                          │
 │             │  authorize/token/registration endpoints (ours)   │
 │             │◄──────────────────────│                          │
 │             │ POST /api/oauth/register (DCR)                    │
 │             │──────────────────────►│  → client_id             │
 │             │◄──────────────────────│                          │
 │             │ GET /api/oauth/authorize?code_challenge=… (PKCE) │
 │             │──────────────────────►│                          │
 │             │        302 to Entra authorize (our PKCE,         │
 │             │        scope=…/.default, NO resource param)      │
 │             │◄──────────────────────│                          │
 │   browser redirect to Entra; user signs in                     │
 │◄───────────────────────────────────────────────────────────────│
 │───────────────────────────────────────────────────────────────►│
 │             │   302 to /api/oauth/callback?code=…&state=…       │
 │             │◄─────────────────────────────────────────────────│
 │             │                       │ POST Entra /token        │
 │             │                       │ (client_secret + PKCE,   │
 │             │                       │  NO resource param)      │
 │             │                       │─────────────────────────►│
 │             │                       │  access + refresh token  │
 │             │                       │◄─────────────────────────│
 │             │                       │ encrypt tokens → session │
 │             │                       │ mint our one-time code   │
 │             │  302 to Claude redirect_uri?code=<ours>&state    │
 │             │◄──────────────────────│                          │
 │             │ POST /api/oauth/token (our code + PKCE verifier)  │
 │             │──────────────────────►│  → opaque Bearer + refresh│
 │             │◄──────────────────────│                          │
 │             │ stores OUR opaque tokens in Claude's user storage│
```

### Tool invocation (every MCP request)

```
Claude.ai                  Our MCP                Microsoft Graph
    │                        │                        │
    │ POST /api/mcp/mcp       │                        │
    │ Authorization: Bearer O │  (O = our opaque token)│
    │────────────────────────►│                        │
    │                        │                        │
    │     verifyMcpToken(O):                           │
    │      - sha256(O) → look up access-token ref      │
    │      - check not expired                         │
    │      - load session (encrypted Graph tokens)     │
    │     → AuthInfo {sessionId, userId, scopes}       │
    │                        │                        │
    │     tool handler called with AuthInfo            │
    │     getGraphTokenForSession(sessionId):          │
    │      - decrypt Graph access token                │
    │      - if within 5 min of expiry, refresh via    │
    │        Entra (client_secret), re-encrypt, store  │
    │                        │                        │
    │                        │ Graph API call         │
    │                        │ Authorization: Bearer G │  (G = real Graph token)
    │                        │───────────────────────►│
    │                        │  result                │
    │                        │◄───────────────────────│
    │                        │                        │
    │  MCP tool result        │                        │
    │◄────────────────────────│                        │
```

When our opaque access token expires (1 hour), Claude uses the refresh token
we issued to get a new pair from `/api/oauth/token` — that maps back to the
same session. Separately, the underlying Graph token is refreshed against
Entra on demand inside `getGraphTokenForSession`, transparently to Claude.

---

## Code layout

```
app/
├── api/
│   ├── mcp/[transport]/route.ts       # MCP HTTP endpoint (POST/GET/DELETE)
│   ├── oauth/
│   │   ├── register/route.ts          # Dynamic Client Registration (RFC 7591)
│   │   ├── authorize/route.ts         # our /authorize → 302 to Entra
│   │   ├── callback/route.ts          # Entra redirect target; mints our code
│   │   └── token/route.ts             # our /token (auth_code + refresh_token)
│   └── health/route.ts                # liveness check
├── .well-known/
│   ├── oauth-protected-resource/      # RFC 9728: points clients at us as AS
│   │   └── route.ts
│   └── oauth-authorization-server/    # RFC 8414: our AS metadata
│       └── route.ts
├── layout.tsx                         # Next.js shell
└── page.tsx                           # public landing page

lib/
├── env.ts                             # env parsing, constants
├── oauth.ts                           # PKCE, token gen, redirect-uri checks
├── store.ts                           # Upstash Redis: clients/sessions/codes,
│                                      #   AES-256-GCM token encryption
├── entra.ts                           # server-side code/refresh exchange w/ Entra
├── auth.ts                            # opaque Bearer → session (verifyMcpToken)
├── graph.ts                           # Graph fetch helpers + per-session token
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

### Why are we our own authorization server (the proxy)?

Two hard Entra constraints forced this design:

1. **No Dynamic Client Registration.** Claude expects to register itself with
   the authorization server at connect time (RFC 7591). Entra has no public
   DCR endpoint, so Claude could never get a `client_id` directly from Entra.
   By being the AS ourselves, we implement DCR and hand Claude a client_id.
2. **Entra rejects `resource`+`scope` together (AADSTS9010010).** When Claude
   drives the flow directly, it sends both a `resource` indicator (RFC 8707)
   and Graph scopes; Entra's v2.0 endpoint rejects that combination
   (`invalid_target`). By taking Claude out of the direct Entra path, *we*
   shape the Entra request — Graph `.default` scope, **no** `resource` param —
   so Entra is happy.

Claude does PKCE against *us*; we do a second, independent PKCE leg against
Entra. Entra tokens never leave our server.

### Why a client secret now?

Because we are the party exchanging authorization codes and refresh tokens
with Entra (server-side), we are a **confidential client** and must
authenticate to Entra's token endpoint with `WCIO_MCP_CLIENT_SECRET`. This is
the opposite of the earlier public-client/PKCE-only design — and it's exactly
what lets us run the code/refresh exchanges that keep Graph tokens off Claude.

### Why drop on-behalf-of (OBO)?

An earlier iteration tried OBO (exchange the user's inbound token for a Graph
token). With the proxy we request Graph delegated scopes **directly** during
the server-side authorization-code exchange, so there is no inbound user token
to exchange and no OBO leg. Simpler, fewer Entra round-trips, and it sidesteps
the same `resource`/audience pitfalls. If we ever need a second downstream API
beyond Graph, we'd request its scopes in the same exchange (or revisit OBO).

---

## Security model

This is now a stateful authorization server that holds Entra tokens, so the
security model is meaningfully different from the old stateless proxy.

### What the server can do
- Issue and validate its own opaque Bearer tokens (mapped to sessions)
- Run authorization-code and refresh-token exchanges with Entra (confidential client)
- Hold Entra access + refresh tokens, **encrypted at rest** (AES-256-GCM in Redis)
- Forward authenticated requests to Graph as the user
- Log non-sensitive metadata (user identifier, tool name, success/failure)

### What the server cannot do
- Act on behalf of a user with no valid session / opaque token
- Access OneDrive/SharePoint/Outlook outside a user's Graph-scoped permissions
  (delegated scopes are gated by the user's effective access)
- Decrypt stored tokens without `WCIO_MCP_TOKEN_ENCRYPTION_KEY`
- Authenticate to Entra without `WCIO_MCP_CLIENT_SECRET`

### Defenses
- **Opaque tokens, not JWTs**: Claude never receives an Entra token. Our Bearer
  is a random 256-bit value; only its sha256 is stored, mapped to a session.
- **PKCE on both legs**: Claude→us and us→Entra each use Authorization Code +
  PKCE (S256). Authorization codes are one-time and expire in 60s.
- **Tokens encrypted at rest**: Entra access/refresh tokens are sealed with
  AES-256-GCM; the key lives only in env, never in Redis or the repo.
- **Short TTLs + rotation**: in-flight auth requests 10 min, opaque access
  tokens 1 h, refresh tokens rotate on use (old one consumed).
- **Redirect-URI validation**: `/authorize` validates the client and its
  redirect URI *before* trusting any redirect; only https or http-loopback
  redirects are accepted at registration.
- **Single-tenant Entra app**: only users from the configured tenant can sign in.

### What attackers can attempt
- **Stealing the opaque Bearer**: limited to its 1-hour window and only grants
  Graph access the user already has; it cannot be replayed against Entra
  directly (it isn't an Entra token).
- **Compromising Redis**: tokens there are encrypted; without the encryption
  key they're useless. Codes/auth-requests are short-lived.
- **Compromising the server process**: higher blast radius than the old
  design (the key and client secret are in env, and live tokens pass through
  memory). Mitigations: encryption at rest, no token logging, short TTLs,
  refresh-token rotation. Rotating the encryption key invalidates all sessions.
- **Forged authorization codes / PKCE downgrade**: blocked by one-time codes,
  S256-only verification (`timingSafeEqual`), and client_id/redirect_uri checks.
- **Prompt injection from file/email content**: the AI client's responsibility;
  the client is expected to confirm with the user before destructive actions.

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
- **Additional downstream APIs beyond Graph**: request their delegated scopes
  in the same server-side authorization-code exchange (or revisit OBO).
- **App-only tokens** for system tasks like scheduled cleanup. Higher risk
  surface; not needed for any current use case.
- **Webhook support**: subscribe to Graph notifications and trigger tools.
  Would require persistent state.
