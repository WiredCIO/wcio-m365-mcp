# Deployment and Rollout

End-to-end checklist from empty repo to working MCP server used by the Wired
CIO team.

---

## Phase 1: GitHub repository

1. In GitHub, create a new **private** repo under the `wiredcio` org:
   - Name: `wcio-m365-mcp`
   - Visibility: Private
   - Don't initialize with README/license — we're pushing local code
2. From the unpacked source folder:
   ```bash
   cd wcio-m365-mcp
   git init
   git add .
   git commit -m "Initial scaffold: Microsoft 365 MCP server"
   git branch -M main
   git remote add origin git@github.com:wiredcio/wcio-m365-mcp.git
   git push -u origin main
   ```

---

## Phase 2: Entra app registration

Follow [ENTRA_SETUP.md](./ENTRA_SETUP.md). At the end you'll have:
- Tenant ID
- Application (client) ID
- Admin-consented permission grants

Don't add the Vercel deployment URL as a redirect URI — only Claude.ai's
callback URL needs to be there.

---

## Phase 3: Vercel deployment

### Option A — via the Vercel dashboard

1. Vercel → **Add New...** → **Project**
2. Import `wiredcio/wcio-m365-mcp` (you may need to grant Vercel access to the org first)
3. **Framework Preset**: Next.js (auto-detected)
4. **Build & Output Settings**: leave defaults
5. **Environment Variables**:
   - `WCIO_MCP_TENANT_ID` = your tenant ID
   - `WCIO_MCP_CLIENT_ID` = your application client ID
   - `WCIO_MCP_BASE_URL` = `https://wcio-m365-mcp.vercel.app` (will be the assigned URL; update after first deploy if Vercel chose a different name)
   - `WCIO_MCP_VERBOSE_LOGGING` = `0` (set to `1` only when debugging)
6. Click **Deploy**.

### Option B — via Vercel CLI

```bash
npm i -g vercel
cd wcio-m365-mcp
vercel link
vercel env add WCIO_MCP_TENANT_ID production
vercel env add WCIO_MCP_CLIENT_ID production
vercel env add WCIO_MCP_BASE_URL production
vercel --prod
```

### Post-deploy verification

After deploy, hit:
```
https://<your-vercel-url>/api/health
```
Expected response:
```json
{
  "status": "ok",
  "name": "wcio-m365-mcp",
  "version": "0.1.0",
  "tenantConfigured": true,
  "baseUrl": "https://...",
  "timestamp": "..."
}
```

And:
```
https://<your-vercel-url>/.well-known/oauth-protected-resource
```
Should return JSON with `authorization_servers` pointing to your tenant's
`login.microsoftonline.com/<tenant-id>/v2.0` URL.

---

## Phase 4: First-user test

You — as the admin — test before rolling out.

1. In Claude.ai → **Settings** → **Connectors** → **Add custom connector**
2. **Name**: `Wired CIO M365`
3. **URL**: `https://<your-vercel-url>/api/mcp/mcp`
4. Save. Claude.ai will:
   - Fetch your /.well-known/oauth-protected-resource
   - Redirect you to Microsoft sign-in
   - You sign in with your @wiredcio.com account
   - First-time consent prompt appears (already pre-consented at admin level, so it shows the granted permissions but no individual approval needed)
   - You're redirected back to Claude.ai
5. Open a new chat. Ask: "Use the Wired CIO M365 connector to list files in my OneDrive root."
6. Claude should call `list_folder` and return your top-level OneDrive items.

### Smoke test the write side

Try the actual Amuze use case:
```
Update Amuze_Data_Dictionary_v5.xlsx in my OneDrive with the version
I'm attaching, using the wcio-m365 update_file tool.
```

Verify in OneDrive that the file is replaced and that version history shows
the prior version is recoverable.

---

## Phase 5: Custom domain (optional)

Once you've validated the Vercel URL works:

1. In Vercel project settings → **Domains** → **Add**
2. Enter `mcp.wiredcio.com`
3. Add the CNAME record Vercel gives you to your DNS
4. Wait for cert provisioning (usually 1–2 minutes)
5. Update `WCIO_MCP_BASE_URL` env var to `https://mcp.wiredcio.com` and redeploy
6. Update the Claude.ai connector URL to `https://mcp.wiredcio.com/api/mcp/mcp`

---

## Phase 6: Org-wide rollout

### Communications

Send an email to the team. Template:

> **Subject**: New Claude integration — Microsoft 365 write access
>
> Team,
>
> Claude can now write to OneDrive, SharePoint, Outlook, Calendar, and Teams
> on your behalf. To enable it:
>
> 1. Go to Claude.ai → Settings → Connectors → Add custom connector
> 2. URL: `https://mcp.wiredcio.com/api/mcp/mcp`
> 3. Sign in with your @wiredcio.com account when prompted
>
> Once connected, Claude can update files in OneDrive (with version history
> preserved), send emails from your account, create calendar invites,
> upload artifacts to SharePoint, and post to Teams chats and channels.
>
> Every action runs under your account — audit logs in Entra and OneDrive
> show your name. Claude will confirm with you before sending anything
> outbound (emails, Teams messages).
>
> Russell

### Audit logging

Microsoft Entra audit logs capture every token issuance under our app:
- Entra admin center → **Identity** → **Monitoring & health** → **Sign-in logs** → filter Application = "Wired CIO M365 MCP"

Microsoft Graph activity logs capture every API call:
- Entra admin center → **Identity** → **Monitoring & health** → **Audit logs**
- For more detail, enable Microsoft Graph Activity Logs in your tenant — these capture every Graph API call by user, endpoint, status

### Revoking access

To remove a user's ability to use the connector:
- Entra → **App registrations** → **Wired CIO M365 MCP** → **Users and groups**
- Or block them at the tenant level via Conditional Access policy

To kill the integration entirely:
- Vercel: pause the project (instant)
- Entra: disable the app registration (revokes all existing tokens within ~1 hour)

---

## Updating the server

Every commit to `main` auto-deploys to production. For sensitive changes (new
scopes, new tool surface), use a feature branch + preview deployment first:

```bash
git checkout -b feature/new-tool
# edit code
git push origin feature/new-tool
```

Vercel will create a preview URL. Test against it by adding it as a separate
custom connector in Claude.ai (use a non-production URL prefix for testing).
Merge to main when ready.

---

## Disaster recovery

The MCP server holds no state. If Vercel disappears:
1. Pull the GitHub repo
2. Deploy elsewhere (Cloudflare Workers, AWS Lambda, Render — anywhere that runs Next.js)
3. Update the Entra app's redirect URI if the new host needs a different one
4. Update `WCIO_MCP_BASE_URL` and the Claude.ai connector URL

Token issuance and validation continue to work as long as the Entra app
exists and Microsoft Graph is reachable.
