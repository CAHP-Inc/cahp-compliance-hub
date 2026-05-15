import type { Role } from './permissions';

/**
 * Email → Role mapping.
 *
 * HARDCODED FOR PR-02. Will be replaced by a SharePoint Users List lookup in PR-04.
 *
 * Rules:
 * - Keys are LOWERCASE email addresses (we normalize on lookup, but keep this consistent)
 * - Only emails in this map can access the application; everyone else sees "Access denied"
 * - Any user in the @newshire.com / @cahphousing.org / etc. M365 tenant who signs in but is
 *   not in this map will be denied — that's the right behavior. M365 tenant membership alone
 *   does not grant CAHP Compliance Hub access.
 *
 * To add or remove team members: edit this file, commit, push. Production picks up the
 * change on the next deploy. Eventually this becomes a SharePoint List you edit through
 * the Settings module — but for now, code is the source of truth.
 */
const EMAIL_ROLE_MAP: Record<string, Role> = {
  // ============================================================
  // Active access list. To add team members: add their email + role here,
  // commit, push. Production picks up the change on the next deploy.
  // ============================================================
  'bturner@newshirepm.com': 'Admin',
  'lheckman@redcedarhomes.com': 'Contributor',

  // Future team members (uncomment and update emails when adding):
  // 'bryan@cahphousing.org': 'Admin',
  // 'stan@vanrock.com': 'Admin',
  // 'cara@newshirepm.com': 'Contributor',
  // 'chris@newshirepm.com': 'Accounting',
};

/**
 * Look up a role for a given email. Returns null if the user is not on the access list.
 */
export function lookupRole(email: string | undefined | null): Role | null {
  if (!email) return null;
  return EMAIL_ROLE_MAP[email.toLowerCase()] ?? null;
}

/**
 * Compute display initials from a name (preferred) or email (fallback).
 * Examples:
 *   "Brandy Turner" → "BT"
 *   "Bryan" → "BR"
 *   "lori@newshire.com" → "LO"
 *   null → "??"
 */
export function getUserInitials(
  name: string | undefined | null,
  email: string | undefined | null
): string {
  if (name) {
    const parts = name.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  }
  if (email) {
    return email.slice(0, 2).toUpperCase();
  }
  return '??';
}

/**
 * Derive a display "org" label from email domain — cosmetic only.
 */
export function extractOrgFromEmail(email: string | undefined | null): string {
  if (!email) return '—';
  const domain = email.split('@')[1] || '';
  const lower = domain.toLowerCase();
  if (lower.includes('newshire')) return 'NewShire';
  if (lower.includes('cahp')) return 'CAHP';
  if (lower.includes('vanrock')) return 'VanRock';
  if (lower.includes('redcedar')) return 'Red Cedar';
  return domain.split('.')[0] || '—';
}
