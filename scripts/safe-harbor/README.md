# Safe Harbor Certification generator

Generates a **SC Code §12-37-220(B)(11)(e) / IRS Rev. Proc. 96-32 Safe Harbor
Certification** letter (modeled on the GSP TTAC template) plus an **Exhibit A**
unit-by-unit AMI analysis, from an AppFolio rent roll.

It also **determines which Rev. Proc. 96-32 set-aside scope** the portfolio
qualifies for:

| Scope | Deep-affordability test | Market cap |
|-------|------------------------|------------|
| **20/50** | ≥ 20% of units at ≤ 50% AMI | ≤ 25% above 80% AMI |
| **40/60** | ≥ 40% of units at ≤ 60% AMI | ≤ 25% above 80% AMI |

(The "≤ 25% above 80%" cap is the same thing as the "≥ 75% low-income" overlay.)

## How units are classified — rent test

Qualification is **rent-based**, which is the basis the certification letter
itself states ("gross rent does not exceed the Maximum Allowable Gross Rent …
under the HUD MTSP methodology"). For each unit:

1. County is detected from the address (Greenville / Spartanburg).
2. Gross rent = contract rent (or market/asking rent for a vacant unit) **plus**
   any tenant-paid utility allowance (`--utility-allowance`, default `$0`).
3. That gross rent is compared to the published Maximum Allowable Gross Rent for
   the unit's county + bedroom size at **50% / 60% / 80% AMI**
   (`hud_limits_fy2026.json`). The deepest tier it fits is its classification.

> **The demographic / income-range report is corroborating only.** AppFolio's
> tenant demographic export is almost entirely "No Income Reported," so it cannot
> classify units. Pass it with `--demographic` and it is summarized in the letter
> as supporting evidence; omit it and the letter is generated from rents alone.

## One-time setup

```powershell
python -m pip install python-docx openpyxl
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force   # only for export_entity.ps1
```

All commands are run from the repo root (`cahp-compliance-hub\`) in PowerShell.

### A. Single LLC

```powershell
# 1. Pull the entity's boilerplate live from the hub (validates the 501(c)(3)
#    exemption chain; writes safe-harbor-output\<slug>.entity.json).
.\scripts\safe-harbor\export_entity.ps1 -EntityTitle "IV SPB II LLC"

# 2. (Optional) open the JSON and fill any blanks it reports
#    (operating-agreement date, certifier relationship, parcels, EIN…).

# 3. Generate the letter + Exhibit A from the AppFolio rent roll.
python scripts\safe-harbor\generate_cert.py `
    --rent-roll   rent_roll-20260611.xlsx `
    --demographic tenant_demographic_06-11-2026-03-15-pm.xlsx `
    --entity      safe-harbor-output\IV_SPB_II_LLC.entity.json `
    --out         safe-harbor-output
```

Without the hub (no PnP access) you can skip step 1: copy
`entity_config.template.json` to `safe-harbor-output\<slug>.entity.json`, fill it
in by hand, and run step 3.

### B. Portfolio / group filing (one parent, many LLCs filed together)

For a group like **IV Fund Global, LLC** that owns multiple single-purpose LLCs
and files for all of them as one group, the tool aggregates every unit across the
sub-LLCs, runs the scope test on the **combined** portfolio, and produces **one**
certification letter with a per-LLC composition table.

```powershell
# 1. Export ONE rent roll per sub-LLC from AppFolio, each filtered to that LLC's
#    Property Group (the "Property Groups: Owner- <LLC>" line in the export tells
#    the tool which units belong to which LLC). Drop them all in one folder, e.g.
#    safe-harbor-output\iv-fund-global\.

# 2. Build a GROUP entity config: copy entity_config.template.json, set
#    portfolio.isGroupFiling = true and portfolio.groupName = "IV Fund Global, LLC".
#    (Shared nonprofit/manager/certification boilerplate applies to the whole group.)

# 3. Pass ALL the rent rolls at once (a folder, or a space-separated list).
python scripts\safe-harbor\generate_cert.py `
    --rent-roll safe-harbor-output\iv-fund-global\ `
    --entity    safe-harbor-output\iv-fund-global.group.json `
    --out       safe-harbor-output
```

You can also list files explicitly instead of a folder:
`--rent-roll roll_spb2.xlsx roll_spb3.xlsx roll_spb4.xlsx`. A group filing is
triggered automatically whenever the rent rolls cover more than one LLC, or
whenever `portfolio.isGroupFiling` is `true`.

### Outputs (either mode)

Land in `safe-harbor-output\` (git-ignored):

- `<name>_Safe_Harbor_Certification_TY2026.docx` (group files get a `_GROUP` tag)
- `<name>_Exhibit_A_Unit_AMI_Analysis.xlsx` — every unit with its Source LLC,
  plus a **Summary** sheet and (for groups) a **Per-LLC** breakdown sheet.

> **Project-level vs group test.** Filing as a group applies the Rev. Proc. 96-32
> set-aside percentages to the *combined* unit pool. The Exhibit's Per-LLC sheet
> also shows how each LLC scores on its own, so you can see whether any single
> LLC would fail standalone even though the portfolio passes (or vice-versa).

## Reading the result

The console (and the letter's Section 3 table) report the four cumulative
percentages and a **Set-Aside Determination**. Units the rent roll can't classify
(missing bedroom count, rent, or county) are flagged **NEEDS REVIEW** in
Exhibit A and counted conservatively: a scope is reported as `QUALIFIES` only if
it passes even when every review unit is treated as failing. Resolve those units
(usually a missing `BD/BA` in AppFolio) and re-run before signing.

## ⚠ Before anyone signs (penalty of perjury)

- **Re-verify the rent limits** in `hud_limits_fy2026.json` against the live SC
  Housing / HUD MTSP source PDFs (links are in the file's `_meta`). The 60% tier
  is derived as `floor(1.2 × 50%)`; confirm against the HUD MTSP 60% rent chart
  if a unit's qualification hinges on the 60% tier.
- **Set the correct utility allowance.** Default is `$0` (compare contract rent
  directly). If tenants pay utilities, pass the allowance so gross rent is tested
  correctly: `--utility-allowance <amount>`.
- **Confirm the exemption chain.** `export_entity.ps1` warns if the property has
  no CAHP-flagged 501(c)(3) member in its ownership — without one the entity does
  not qualify under §12-37-220(B)(11)(e).

## Files

| File | Committed? | Purpose |
|------|-----------|---------|
| `generate_cert.py` | yes | The generator (no PII). |
| `export_entity.ps1` | yes | Pulls entity facts from the hub via PnP. |
| `hud_limits_fy2026.json` | yes | FY2026 rent limits (public data) + sources. |
| `entity_config.template.json` | yes | Placeholder config to copy/fill by hand. |
| `safe-harbor-output\*` | **no — git-ignored** | Real entity configs, rent rolls, and generated letters (tenant + entity PII). |
