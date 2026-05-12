# Scripts

PowerShell utilities for one-time / operational tasks. Run from the project root.

---

## `provision-sharepoint.ps1` — PR-03

Provisions all SharePoint Lists and Document Libraries for the CAHP Compliance Hub.

### Prerequisites

1. A SharePoint site to host the lists (Team site recommended, private)
2. Site Owner or Site Collection Admin permissions on that site
3. PnP.PowerShell module installed:
   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
   ```

### Run it

```powershell
cd C:\Users\brand\code\cahp-compliance-hub
.\scripts\provision-sharepoint.ps1 -SiteUrl "https://newshirepm.sharepoint.com/sites/CAHPHub"
```

A browser window opens for sign-in (one time, then cached). The script then runs in ~60-90 seconds and prints:

- ✓ for each newly created list/column
- → for items that already exist (skipped, no-op)
- ! for any failures (logged but script continues)

### What it creates

**14 SharePoint Lists** (data layer):
Users, Owners, Properties, Ownership, Submittals, Correspondence, Communications, Outstanding, Documents, Billing, Disbursements, AuditLog, Notifications, Notes

**13 Document Libraries** (file storage):
Operating Agreements, LURAs, AMI Certifications, Rent Rolls, Insurance Certificates, DOR Correspondence Files, Owner Communication Files, Entity Formation, Governance, Submittal Packages, Tax Certificates, Compliance Documents, Backups

Each library gets a `PropertyID` metadata column and an `ExpirationDate` column for filtering and renewal tracking.

### After running

1. **Verify in SharePoint** — visit `<site-url>/_layouts/15/viewlsts.aspx` to see all lists and libraries
2. **Note the Graph site ID** — the script prints it at the end, in `hostname:/path` format. You need this for the next step.
3. **Set GitHub Actions Variable:**
   - Name: `VITE_SHAREPOINT_SITE`
   - Value: the Graph site ID (e.g., `newshirepm.sharepoint.com:/sites/CAHPHub`)

### Idempotent

The script is safe to re-run. New columns added in code will be added on re-run; existing items are detected and skipped.

### Adding a new column later

1. Find the relevant `Ensure-CAHPList` call in the script
2. Add a new hashtable entry to the `Columns` array
3. Re-run the script — only the new column is added

---

## Future scripts

- `seed-data.ps1` — populate the Users list and initial seed records (Phase 1 polish)
- `lock-audit-log.ps1` — apply write-once permissions to the AuditLog list (Phase 4 / PR-07)
- `backup-export.ps1` — manual full-database export (the Power Automate equivalent)
