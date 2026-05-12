# CAHP Compliance Hub

Internal application for managing CAHP (Carolina Affordable Housing Project) property tax abatement filings across the Carolinas.

Authored by Brandy Turner — NewShire Property Management.

---

## Status

**Phase 1 / PR-01: Scaffolding** — application shell, navigation drawer, design tokens, deployment pipeline. No data layer yet. See `Phase 1 Build Progress` on the My Day landing page for the full PR sequence.

---

## Tech stack

- **React 18** + **TypeScript** + **Vite**
- **Tailwind CSS** with NewShire design tokens
- **react-router-dom** for client-side routing
- **MSAL** (wired in PR-02) for M365 authentication
- **Microsoft Graph SDK** (wired in PR-04) for SharePoint List CRUD
- **GitHub Pages** for hosting via GitHub Actions

---

## PR-01 Setup Checklist

Run through this once. Total time: ~15 minutes.

### 1. Create the GitHub repo

```
Repo name: cahp-compliance-hub
Visibility: Private
Initialize: empty (no README, no .gitignore)
```

### 2. Unzip this package into the new repo

```bash
cd ~/code
unzip cahp-compliance-hub.zip
cd cahp-compliance-hub
git init
git remote add origin https://github.com/<your-username>/cahp-compliance-hub.git
```

### 3. Install dependencies and verify it runs locally

```bash
npm install
npm run dev
```

Open the printed `http://localhost:5173` URL. You should see:
- Teal header with hamburger ☰
- My Day landing page with PR-01 banner
- Phase 1 Build Progress card
- Dev Role Switcher card

Click the hamburger and verify the drawer slides in with 6 grouped sections. Switch roles in the Dev Role Switcher and reopen the drawer to confirm Settings disappears for Contributor and most groups disappear for Accounting.

### 4. Initial commit and push

```bash
git add .
git commit -m "PR-01: Scaffolding + AppShell + GH Pages deploy"
git branch -M main
git push -u origin main
```

### 5. Enable GitHub Pages

In the repo on GitHub:

```
Settings → Pages
Source: GitHub Actions
```

### 6. Set the base path repository variable

The GH Pages site lives at `https://<username>.github.io/cahp-compliance-hub/` — Vite needs to know about the `/cahp-compliance-hub/` subpath. Set this as a repo variable so the workflow picks it up:

```
Settings → Secrets and variables → Actions → Variables tab → New repository variable
Name: VITE_BASE_PATH
Value: /cahp-compliance-hub/
```

### 7. Trigger the deploy

```bash
git commit --allow-empty -m "Trigger deploy"
git push
```

Watch the deploy run in the **Actions** tab. When green (~2 minutes), open `https://<username>.github.io/cahp-compliance-hub/` — same UI as local.

---

## Scripts

```bash
npm run dev         # Start dev server on http://localhost:5173
npm run build       # Production build to ./dist
npm run preview     # Preview the production build locally
npm run typecheck   # Run TypeScript without emitting
```

---

## Project structure

```
cahp-compliance-hub/
├── .github/workflows/deploy.yml    # GH Pages CI/CD
├── public/favicon.svg              # CAHP "C" mark
├── src/
│   ├── components/
│   │   ├── layout/                 # AppShell, Header, SidebarNav
│   │   └── ui/                     # Icon (more in PR-05+)
│   ├── lib/
│   │   ├── permissions.ts          # Roles + canView/canDo
│   │   └── session.tsx             # Session context (stub → MSAL in PR-02)
│   ├── pages/
│   │   ├── MyDay.tsx               # Landing page
│   │   └── PlaceholderPage.tsx     # Used until each module is built
│   ├── styles/globals.css          # Design tokens + Tailwind
│   ├── App.tsx                     # Router
│   ├── main.tsx                    # Entry point
│   └── vite-env.d.ts               # Env var types
├── .env.example                    # Copy to .env.local and fill in
├── index.html
├── package.json
├── tailwind.config.js              # NewShire palette
├── tsconfig.json
└── vite.config.ts
```

---

## What ships in PR-01

- Application shell (header, hamburger drawer, content area)
- 6-group role-filtered navigation
- NewShire design tokens (teal/gold palette, fonts, animations)
- Routes for all 16 modules (15 wired to placeholders, My Day fully styled)
- Dev role switcher to preview Contributor / Accounting views
- TypeScript-strict mode
- GitHub Actions deploy to GH Pages

---

## What does NOT ship in PR-01

- Real auth (MSAL stub — `session.tsx` returns a static "Brandy / Admin" user)
- Any data layer (no Graph SDK, no SharePoint connection)
- Any module content (everything except My Day is a placeholder)
- Search bar (Phase 2)
- Notifications (Phase 2)
- Untagged Docs badge counts (PR-05 when documents wire up)

---

## Next: PR-02

PR-02 wires real M365 authentication. Before starting PR-02:

1. Register an Azure AD app (see `.env.example` for the exact steps)
2. Copy `.env.example` to `.env.local`
3. Fill in `VITE_AZURE_CLIENT_ID` and `VITE_AZURE_TENANT_ID`

When ready, ask Claude for PR-02.

---

## License

Internal use only. © 2026 Brandy Turner / NewShire Property Management.
