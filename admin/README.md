# EVAA Admin Portal (web)

Static admin tool that signs portal admins in via MSAL.js and talks to Microsoft Graph directly. Lives at `https://markmeevaa.github.io/directory/admin/` once published.

Phase 1 scope (this commit): **read-only**. Sign in → verify member of `EVAA Portal Admins` group → list managed EVAA/Fusion Unified groups → drill into a group to see its directors (owners) and members.

Phase 2 (next): add/remove members, add/remove owners, jobTitle edits, audit log.

## One-time Azure AD setup (required before sign-in works)

The existing EVAA app registration (`74b79a84-6ea0-4c32-beae-25ed9bdc249f`) needs a **redirect URI** and a couple of **delegated Graph scopes** added so user sign-in works.

1. Go to **portal.azure.com** → search "App registrations" → open the EVAA app (`74b79a84-...`).
2. Left rail → **Authentication**:
   - Under **Platform configurations**, click **+ Add a platform** → **Single-page application**.
   - Redirect URI: `https://markmeevaa.github.io/directory/admin/`
   - Leave Access tokens / ID tokens unchecked (we use MSAL.js v3 PKCE).
   - **Save**.
3. Left rail → **API permissions**:
   - Click **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**.
   - Tick:
     - `User.Read` (probably already present)
     - `Group.Read.All`
     - `GroupMember.Read.All`
   - **Add permissions**.
   - Then click **Grant admin consent for Eastview Athletic Association**.

That's it. The web app should then work for anyone in the `EVAA Portal Admins` group.

For local-testing **only** (optional): also add `http://localhost:8000/admin/` as an SPA redirect URI if you want to run a local web server while iterating.

## Local dev / preview

The app is plain HTML/JS — no build step. To preview locally:

```bash
cd /path/to/directory
python3 -m http.server 8000
# Then visit http://localhost:8000/admin/
```

Sign-in will fail unless `http://localhost:8000/admin/` is also registered as a redirect URI on the app registration (see optional step above).

## File layout

```
admin/
├── index.html       Markup + the three views (signin / groups / group-detail)
├── css/styles.css   EVAA-blue branded styles, matches directory pages
├── js/auth.js       MSAL config + sign-in/sign-out + token acquisition
├── js/graph.js      Microsoft Graph fetch wrapper (paginated)
├── js/app.js        Main app: state, rendering, click handlers
└── README.md        This file
```

## How it relates to the rest of the EVAA stack

- **Reads** group/member/owner data directly from Microsoft Graph (delegated permissions, user-scoped). No SharePoint dependency for reads.
- For Phase 2 writes (add/remove member), the plan is to POST to the same `MemberRequests` SharePoint list that the existing Power Apps canvas portal writes to. That preserves the entire **EVAA MemberRequest Approval Flow** automation unchanged — flow doesn't care which UI created the request.
- The existing **canvas app stays running for sport presidents** until/unless we migrate them too (Phase 2/3). This admin tool sits alongside it, not on top of it.

## Deploy

GitHub Pages serves whatever's on `main`. `git push origin main` → live in ~1 minute at `https://markmeevaa.github.io/directory/admin/`. No build step, no CI.
