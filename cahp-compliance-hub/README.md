# CAHP Compliance Hub

Internal application for managing CAHP (Carolina Affordable Housing Project) property tax abatement filings across the Carolinas.

Authored by Brandy Turner — NewShire Property Management.

---

## Status

**Phase 1 / PR-02: M365 Authentication** — application is now gated behind Microsoft 365 sign-in. Role is determined by email mapping. Next: PR-03 provisions the SharePoint backend.

---

## Tech stack

- **React 18** + **TypeScript** + **Vite**
- **Tailwind CSS** with NewShire design tokens
- **react-router-dom** for client-side routing
- **MSAL** (`@azure/msal-browser` + `@azure/msal-react`) for M365 authentication
- **Microsoft Graph SDK** (wired in PR-04) for SharePoint List CRUD
- **GitHub Pages** for hosting via GitHub Actions

---

## PR-02 Setup Checklist

Before running PR-02 locally or deploying it, you need three things in place:

### 1. Azure AD app registration (one-time, ~10 min)

In [portal.azure.com](https://portal.azure.com):

1. Microsoft Entra ID → App registrations → **+ New registration**
2. Name: `CAHP Compliance Hub`
3. Supported account types: *Accounts in this organizational directory only (Single tenant)*
4. Redirect URI: Single-page application (SPA) → `http://localhost:5173/`
5. Click **Register**
6. From the Overview page, copy:
   - **Application (client) ID**
   - **Directory (tenant) ID**
7. Authentication → SPA section → Add URI: `https://cahp-inc.github.io/cahp-compliance-hub/` → Save
8. API permissions → Add Microsoft Graph delegated: `User.Read`, `Sites.Read.All`, `Sites.ReadWrite.All`
9. Click **Grant admin consent for [tenant]** — all three should show ✓ Granted

### 2. Edit `src/lib/roleMap.ts` with your real email

Open the file. There's a single placeholder entry — replace it with your email:

```ts
const EMAIL_ROLE_MAP: Record<string, Role> = {
  'brandy.turner@newshire.com': 'Admin',  // ← YOUR REAL EMAIL HERE, lowercase
};
```

Use the exact email your M365 account signs in with. Without this, the sign-in flow will succeed but you'll hit "Access denied."

### 3. Create `.env.local` with the Azure values

Copy `.env.example` to `.env.local` and fill in the Azure values from step 1:

```bash
# .env.local
VITE_AZURE_CLIENT_ID=<your-client-id-from-azure>
VITE_AZURE_TENANT_ID=<your-tenant-id-from-azure>
```

`.env.local` is gitignored — never commit it. Each developer maintains their own.

---

## Running locally

```bash
npm install      # one-time
npm run dev      # starts on http://localhost:5173/
```

You should see the sign-in screen. Click **Sign in with Microsoft** → you'll redirect to Microsoft, authenticate, and come back signed in.

If your email is in the role map, you'll land on My Day. If not, you'll see "Access denied" — edit `roleMap.ts`, save, and refresh.

---

## Deploying

PR-01 already wired the GitHub Actions deployment. Add the Azure env vars as repo variables:

1. Repo → Settings → Secrets and variables → **Actions** → Variables tab
2. Add two new variables:
   - `VITE_AZURE_CLIENT_ID` = (your client ID)
   - `VITE_AZURE_TENANT_ID` = (your tenant ID)
3. Push a commit (or empty commit `git commit --allow-empty -m "Trigger deploy"`)
4. Wait for the Actions workflow to complete
5. Visit the production URL — sign-in screen should appear

---

## Scripts

```bash
npm run dev         # Dev server on http://localhost:5173
npm run build       # Production build to ./dist
npm run preview     # Preview the production build locally
npm run typecheck   # TypeScript without emitting
```

---

## Project structure

```
cahp-compliance-hub/
├── .github/workflows/deploy.yml    # GH Pages CI/CD
├── public/favicon.svg
├── src/
│   ├── components/
│   │   ├── auth/SignInGate.tsx     # NEW: sign-in / access-denied gate
│   │   ├── layout/
│   │   │   ├── AppShell.tsx
│   │   │   ├── Header.tsx          # MODIFIED: uses UserMenu
│   │   │   ├── SidebarNav.tsx
│   │   │   └── UserMenu.tsx        # NEW: avatar dropdown with sign out
│   │   └── ui/Icon.tsx
│   ├── lib/
│   │   ├── auth/
│   │   │   └── msalConfig.ts       # NEW: MSAL setup
│   │   ├── permissions.ts          # Roles + canView/canDo
│   │   ├── roleMap.ts              # NEW: email → role (edit this file)
│   │   └── session.tsx             # MODIFIED: real MSAL-backed session
│   ├── pages/
│   │   ├── MyDay.tsx               # MODIFIED: dev-only role override
│   │   └── PlaceholderPage.tsx
│   ├── styles/globals.css
│   ├── App.tsx                     # MODIFIED: wrapped in SignInGate
│   ├── main.tsx                    # MODIFIED: MSAL bootstrap
│   └── vite-env.d.ts
├── .env.example
├── .gitignore
├── index.html
├── package.json
├── tailwind.config.js
├── tsconfig.json
└── vite.config.ts
```

---

## What ships in PR-02

- Microsoft 365 sign-in via MSAL redirect flow
- Branded sign-in screen and access-denied screen
- User menu dropdown from avatar (name, email, role, sign out)
- Role determined automatically from `roleMap.ts` (email → role)
- Dev-only "View As" role override on My Day (hidden in production builds)
- All nav and permission gating now driven by the real signed-in user

---

## What does NOT ship in PR-02

- SharePoint data layer (PR-03 provisions lists, PR-04 wires Graph SDK)
- Module content (Properties, Owners, etc. still placeholders)
- Email-based notifications (Phase 2)
- Search bar (Phase 2)
- Users list editable through the UI — for now, edit `roleMap.ts` in code

---

## Adding team members

For PR-02, adding a user is a code change:

1. Open `src/lib/roleMap.ts`
2. Add a new line to `EMAIL_ROLE_MAP` with their lowercase email and role
3. Commit, push, wait for deploy
4. Tell them to sign in

In PR-04+, this becomes a SharePoint Users List you edit through the Settings module.

---

## License

Internal use only. © 2026 Brandy Turner / NewShire Property Management.
