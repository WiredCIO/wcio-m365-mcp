# Entra ID App Registration

Step-by-step setup for the Wired CIO M365 MCP server in Microsoft Entra ID
(formerly Azure AD).

Total time: ~15 minutes.
Required role: Application Administrator or Global Administrator in your Entra tenant.

---

## 1. Create the app registration

1. Go to [Entra admin center](https://entra.microsoft.com/) → **Identity** → **Applications** → **App registrations**.
2. Click **New registration**.
3. Fill in:
   - **Name**: `Wired CIO M365 MCP`
   - **Supported account types**: **Accounts in this organizational directory only (Wired CIO only — single tenant)**
   - **Redirect URI** — leave blank for now. You'll add the MCP server's own callback URL (`<base>/api/oauth/callback`) once it's deployed. (This server fronts Entra with its own OAuth authorization server, so Entra redirects back to *us*, not directly to Claude.)
4. Click **Register**.

Save these values from the Overview page — you'll need them for Vercel env vars:
- **Application (client) ID**
- **Directory (tenant) ID**

---

## 2. Configure API permissions

The MCP server uses delegated Microsoft Graph permissions so every action runs under the calling user's identity.

1. In the app registration, open **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Add these permissions:

   **Files & SharePoint (Option B baseline)**
   - `Files.ReadWrite.All` — read and write all files the user can access
   - `Sites.ReadWrite.All` — read and write items in all SharePoint sites the user can access

   **Outlook Mail**
   - `Mail.Send` — send mail as the user
   - `Mail.ReadWrite` — read, create, and modify mail (drafts, replies, forward)

   **Calendar**
   - `Calendars.ReadWrite` — read and write the user's calendar
   - `OnlineMeetings.ReadWrite` — required to create Teams meeting links via create_event

   **Teams**
   - `Chat.ReadWrite` — read and post into the user's chats
   - `ChannelMessage.Send` — post into team channels the user is a member of
   - `ChannelMessage.ReadWrite` — required for replying to channel messages

   **Identity**
   - `User.Read` — sign-in and read user profile (default, usually already present)
   - `offline_access` — refresh tokens, so users don't need to re-authenticate every hour
   - `openid` — required for OAuth 2.1 flows

3. Click **Grant admin consent for Wired CIO**.
   - Confirm the prompt.
   - Each permission row should switch to "Granted for Wired CIO" with a green check.

Once consented, every user in the tenant can authenticate without their own consent prompt.

---

## 3. Authentication settings

1. In the app, open **Authentication** → **Add a platform** → **Web**.
2. **Redirect URIs** — add **this MCP server's** callback (NOT Claude's). It is your base URL plus `/api/oauth/callback`, e.g.:
   ```
   https://wcio-m365-mcp.vercel.app/api/oauth/callback
   ```
   Entra redirects here after the user signs in; the server then completes the
   flow back to Claude. If you use a custom domain, register that domain's
   callback instead (and update `WCIO_MCP_BASE_URL` to match).

3. **Front-channel logout URL** — leave blank.
4. **Implicit grant and hybrid flows** — leave both unchecked. We use authorization code with PKCE only.
5. Under **Advanced settings**:
   - **Allow public client flows**: **No** (this is a confidential client — it authenticates to Entra with a client secret).
6. Click **Configure / Save**.

> You do **not** need to configure "Expose an API" / an Application ID URI. The
> server requests Microsoft Graph scopes directly during the server-side
> authorization-code exchange; there is no custom API audience to expose.

---

## 4. Token configuration (optional but recommended)

1. Open **Token configuration** → **Add optional claim**.
2. Token type: **Access token**.
3. Add: `email`, `upn`, `preferred_username`.
4. Save. These claims make audit logs more readable but aren't required for token validation.

---

## 5. Branding (optional)

Under **Branding & properties**:
- Upload the Wired CIO horizontal logo (light or dark background variant works — Entra uses it on the sign-in consent page)
- Set the home page URL to `https://wiredcio.com`
- Set the terms of service / privacy policy URLs if you have them

This makes the consent screen look professional when users first authenticate.

---

## 6. Capture values for Vercel

Set these in Vercel → Project → Settings → Environment Variables:

| Vercel env var | Where to find it |
|---|---|
| `WCIO_MCP_TENANT_ID` | App registration → Overview → Directory (tenant) ID |
| `WCIO_MCP_CLIENT_ID` | App registration → Overview → Application (client) ID |
| `WCIO_MCP_CLIENT_SECRET` | App registration → Certificates & secrets → New client secret → copy the **Value** |
| `WCIO_MCP_BASE_URL` | The Vercel deployment URL once you deploy (next doc) |
| `WCIO_MCP_TOKEN_ENCRYPTION_KEY` | Generate locally: `openssl rand -base64 32` |

A **client secret is required** — this server is a confidential OAuth client. It
runs the authorization-code and refresh-token exchanges with Entra server-side
(Claude never talks to Entra directly), which is what lets us avoid Entra's
`resource`+`scope` rejection (AADSTS9010010) and work around Entra's lack of
dynamic client registration.

You also need a **Redis store** for OAuth state, sessions, and (encrypted)
refresh tokens: in Vercel go to **Storage → add a Redis (Upstash) integration**
and connect it to this project. It injects the `KV_REST_API_URL` /
`KV_REST_API_TOKEN` env vars automatically.

---

## Common gotchas

**"AADSTS500113: No reply address is registered"** / **"AADSTS50011: redirect URI mismatch"**
The redirect URI in the app registration doesn't exactly match `<WCIO_MCP_BASE_URL>/api/oauth/callback`. Confirm the value character-for-character (scheme, host, path, no stray trailing slash) and that `WCIO_MCP_BASE_URL` matches your actual deployment URL.

**"AADSTS9010010: resource parameter doesn't match the requested scopes"**
This is the failure this proxy design exists to prevent. If you see it, something is sending Claude directly to Entra again — confirm `/.well-known/oauth-protected-resource` lists **this server's** base URL under `authorization_servers` (not `login.microsoftonline.com`), and that the connector in Claude was re-added after deploying.

**"AADSTS65001: The user or administrator has not consented"**
Admin consent wasn't granted, or was granted before all required permissions were added. Re-run **Grant admin consent for Wired CIO** in the API permissions blade.

**"invalid_grant" / "offline_access" errors at the callback**
The server needs a refresh token from Entra. Ensure `offline_access` is in the granted delegated permissions (section 2) and admin consent has been re-run.

**Graph calls return 403**
The user account doesn't have access to the resource being requested (e.g. a SharePoint site they aren't a member of). Delegated permissions are gated by the user's effective access, not just the app's declared permissions.

**Connecting fails immediately with a 500 before any sign-in prompt**
Usually a missing env var. All of `WCIO_MCP_CLIENT_SECRET`, `WCIO_MCP_TOKEN_ENCRYPTION_KEY`, and the Redis integration vars must be set — the server validates them at boot and every route fails fast if any are absent.
