# Owner Intake Form — Design Note (Build Later)

**Status:** Idea / not started. Captured 2026-05-22.

## Problem

Onboarding a new owner today means our team types every field (legal name, type, state, EIN, contact email, members, etc.) into the hub by hand from whatever the owner sent over email. That's slow, error-prone, and forces back-and-forth when fields are missing.

## Goal

Let prospective owners submit their own initial information via a form, land it in a staging area, and let our team review/approve it — at which point it seeds the real `Owners` list (and related rows) with no re-typing.

## Chosen approach — Option 1: Microsoft Forms → staging list → review queue in the hub

Why this option (vs. a branded GitHub-Pages form + Power Automate, or an Azure Function backend):
- Owners don't have M365 accounts, so they can't write to SharePoint through the hub's normal MSAL+Graph path. We need a public surface.
- Forms is free, requires no licensing, and owners need no account.
- The heavy lift — turning a submission into a clean `Owners` + `Ownership` + contact set — lives in the hub, where the validation logic from [OwnerNew.tsx](../src/pages/OwnerNew.tsx) can be reused.
- If brand polish on the form itself becomes important later, the front end can be swapped for a branded GitHub-Pages page + Power Automate HTTP trigger **without changing the approval-side code** — the staging list contract is the seam.

### Architecture

```
Owner -> Microsoft Forms (public link)
       -> Power Automate (Forms response trigger)
       -> SharePoint list: OwnerIntake (one row per submission, Status="Pending")
       -> Hub: "Pending Intakes" review page
            -> Team reviews/edits/fixes the data
            -> "Approve" -> creates Owners row (+ Ownership, Contacts) -> marks intake row Status="Approved"
            -> "Reject"  -> marks intake row Status="Rejected" with reason
```

### `OwnerIntake` SharePoint list — proposed schema

Mirrors the `Owners` shape so approval is a near-1:1 copy. Extra columns capture intake metadata.

| Field | Type | Notes |
|---|---|---|
| `Title` | Text | Legal name (matches `Owners.Title`) |
| `OwnerType` | Choice | Individual / LLC / Nonprofit / Trust / Corporation / Limited Partnership / General Partnership |
| `OwnerState` | Text | State of formation/residence |
| `TaxID` | Text | EIN or SSN — masked in the review UI like `OwnerDetail` does today |
| `ContactEmail` | Text | Primary email |
| `ContactPhone` | Text | Primary phone (not on `Owners` yet — could prompt adding it there) |
| `ContactName` | Text | Person submitting, if different from the entity |
| `MailingAddress` | Multi-line text | Street / City / State / Zip — single textarea is fine for intake; team normalizes on approval |
| `OwnerNotes` | Multi-line text | Free-form "anything else we should know" |
| `MembersJSON` | Multi-line text | Optional. JSON array of `{ name, role, percent }` for LLC/Nonprofit owners. Approval step turns these into `Ownership` rows; members the owner names must be matched to existing `Owners` records or created on the fly during approval. |
| `SubmittedAt` | DateTime | Auto-set by Power Automate from the Forms response timestamp |
| `SubmittedByEmail` | Text | From the Forms response (Forms can capture email even without sign-in if we ask) |
| `Status` | Choice | `Pending` / `Approved` / `Rejected` (default Pending) |
| `ReviewedByEmail` | Text | Set on approve/reject |
| `ReviewedAt` | DateTime | Set on approve/reject |
| `RejectionReason` | Multi-line text | Set on reject |
| `LinkedOwnerLookupId` | Lookup → Owners | Set on approve — points to the `Owners` row that was created |

### What the Forms questions ask (owner-facing)

Goal: only collect what the owner uniquely knows. Anything our team can look up (county filings, prior tax IDs we already have) is left off.

1. Owner type (single-select, plain-language wording — not our enum labels)
2. Legal name
3. State of formation (or state of residence for individuals)
4. EIN (or SSN for individuals) — optional, "we can collect this securely later if you prefer"
5. Primary contact email
6. Primary contact phone
7. Primary contact name (if different from the entity)
8. Mailing address (single multi-line — we normalize)
9. For LLCs / Nonprofits / Trusts / LPs only (Forms branching): "List members/partners/trustees with their role and ownership %" — free-text or one repeating section per member. Document that we'll confirm each member during onboarding.
10. Anything else we should know? (free text)
11. Consent / acknowledgement line (optional)

### What the hub's "Pending Intakes" review page does

New route, e.g. `/intakes` (super-admin / compliance-team only — gate via `roleMap.ts`).
Lists `OwnerIntake` rows with `Status === 'Pending'`, newest first. Each row opens a detail view that:

- Renders the submitted data side-by-side with a pre-filled version of the **existing** `OwnerNew` wizard fields, so the reviewer can edit/normalize before approving.
- Validates with the same rules as `OwnerNew.tsx` (legal name required, member sum-to-100% warning, etc.) — extract those into a shared helper so both pages share validation.
- "Approve" button calls a single hub-side function that:
  1. Creates the `Owners` row with the edited values.
  2. For each parsed member in `MembersJSON`: either links to an existing `Owners` record (match-by-name UI) or creates a new one inline, then creates the `Ownership` row.
  3. Creates any contact rows the intake produced.
  4. Updates the `OwnerIntake` row: `Status='Approved'`, `LinkedOwnerLookupId`, `ReviewedByEmail`, `ReviewedAt`.
- "Reject" prompts for a reason, sets `Status='Rejected'`, `RejectionReason`, `ReviewedByEmail`, `ReviewedAt`.

### Power Automate flow

One flow, owned by a service account or by Brandy:

- Trigger: "When a new response is submitted" (Microsoft Forms)
- Action: "Get response details"
- Action: "Create item" in `OwnerIntake` with the mapped fields, `Status='Pending'`, `SubmittedAt=utcNow()`
- (Optional) Action: send a Teams/email notification to the compliance team so they know an intake is waiting

No premium connectors required — the Forms trigger + SharePoint create-item are both standard.

## Open questions to resolve before building

- Do we want one Forms form covering all owner types (with branching) or separate forms per type (Individual vs entity)? Branching is fewer URLs to manage but a longer single form.
- Do we want to capture document uploads in the intake (e.g., a redacted EIN letter)? Forms supports file upload but only for signed-in respondents within the tenant — likely a no for external owners. Defer document collection to post-approval.
- Should rejected intakes be hard-deleted after N days, or kept forever for audit?
- Notification channel for new pending intakes — Teams channel, individual emails, or just rely on the hub's existing notification bell?

## Not in scope for v1

- Branded form styling (would mean moving off Forms — see the alternative options in the original discussion).
- Self-service status checking for the submitter (they can email and ask).
- Auto-approval / any kind of "trust this submitter" shortcut — every intake gets a human review.
