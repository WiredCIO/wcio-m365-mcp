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
   - **Redirect URI** — leave blank for now. You'll add Claude.ai's callback URL once you have the MCP server deployed.
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
2. **Redirect URIs** — add Claude.ai's callback. The current value for claude.ai is:
   ```
   https://claude.ai/api/mcp/auth_callback
   ```
   Confirm the current Claude callback URL via [Anthropic's MCP docs](https://docs.claude.com/) before saving — Anthropic occasionally adds new ones for new regions or product surfaces.

3. **Front-channel logout URL** — leave blank.
4. **Implicit grant and hybrid flows** — leave both unchecked. We use authorization code with PKCE only.
5. Under **Advanced settings**:
   - **Allow public client flows**: **No**
6. Click **Configure / Save**.

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

You'll need three values to set in Vercel environment variables:

| Vercel env var | Where to find it |
|---|---|
| `WCIO_MCP_TENANT_ID` | App registration → Overview → Directory (tenant) ID |
| `WCIO_MCP_CLIENT_ID` | App registration → Overview → Application (client) ID |
| `WCIO_MCP_BASE_URL` | The Vercel deployment URL once you deploy (next doc) |

No client secret is needed — the MCP server is a public OAuth client that uses PKCE.

---

## Common gotchas

**"AADSTS500113: No reply address is registered"**
The redirect URI in the app registration doesn't match what Claude.ai is sending. Confirm the callback URL value exactly, including trailing slashes or lack thereof.

**"AADSTS65001: The user or administrator has not consented"**
Admin consent wasn't granted, or was granted before all required permissions were added. Re-run **Grant admin consent for Wired CIO** in the API permissions blade.

**Tokens validate but Graph calls return 403**
The user account doesn't have access to the resource being requested (e.g. a SharePoint site they aren't a member of). Delegated permissions are gated by the user's effective access, not just the app's declared permissions.

**Tokens validate but the MCP server rejects them**
Check that `aud` claim equals `https://graph.microsoft.com` or `00000003-0000-0000-c000-000000000000`. If you see a custom audience, the OAuth flow isn't requesting Graph scopes correctly — likely a scope formatting issue in Claude.ai's request.
