import type { Role } from './permissions';
import { getListItems, LIST_NAMES, type AccessListEntry } from './sharepoint';

/**
 * Access control: who can sign in and at what role.
 *
 * Primary source of truth: the `Access List` SharePoint list. Managed
 * in-app at Settings → Access List. Loaded once per session (cached
 * here) and refreshable when the editor saves.
 *
 * Hardcoded fallback: the FALLBACK_EMAIL_ROLE_MAP below. Used if the
 * Access List doesn't exist yet (pre-provisioning), can't be reached,
 * or returns no rows. Keeps the original team from getting locked out
 * during outages or before the list is set up.
 */

const FALLBACK_EMAIL_ROLE_MAP: Record<string, Role> = {
  'bturner@newshirepm.com': 'Admin',
  'stan@newshirepm.com': 'Admin',
  'bdebruin@redcedarhomes.com': 'Admin',
  'lheckman@redcedarhomes.com': 'Contributor',
};

const FALLBACK_TEAM_MEMBERS: TeamMember[] = [
  { email: 'bturner@newshirepm.com',     name: 'Brandy Turner',  role: 'Admin' },
  { email: 'stan@newshirepm.com',        name: 'Stan',           role: 'Admin' },
  { email: 'bdebruin@redcedarhomes.com', name: 'Bryan DeBruin',  role: 'Admin' },
  { email: 'lheckman@redcedarhomes.com', name: 'Lori Heckman',   role: 'Contributor' },
];

export interface TeamMember {
  email: string;
  name: string;
  role: Role;
  org?: string;
}

// In-memory cache, populated by loadAccessList()
let cachedEmailRoleMap: Record<string, Role> | null = null;
let cachedTeamMembers: TeamMember[] | null = null;
let cachedSource: 'sharepoint' | 'fallback' | 'unloaded' = 'unloaded';

/**
 * Fetch the Access List from SharePoint and warm the cache. Falls back
 * to the hardcoded map on any error (list missing, network, permissions).
 *
 * Called once at session bootstrap (see lib/session.tsx). Re-runnable —
 * the editor calls it after a save so updates propagate.
 */
export async function loadAccessList(): Promise<void> {
  try {
    const rows = await getListItems<AccessListEntry>(LIST_NAMES.AccessList, { top: 500 });
    if (rows.length === 0) {
      // List exists but is empty — keep the fallback active so we don't
      // accidentally lock everyone out during initial setup.
      cachedEmailRoleMap = { ...FALLBACK_EMAIL_ROLE_MAP };
      cachedTeamMembers = [...FALLBACK_TEAM_MEMBERS];
      cachedSource = 'fallback';
      return;
    }
    const map: Record<string, Role> = {};
    const members: TeamMember[] = [];
    for (const row of rows) {
      const email = (row.fields.Title ?? '').trim().toLowerCase();
      const role = row.fields.AccessRole;
      // AccessActive defaults to true if undefined (back-compat with rows
      // created before the column existed). Explicit `false` denies access.
      const active = row.fields.AccessActive !== false;
      if (!email || !role || !active) continue;
      map[email] = role;
      members.push({
        email,
        name: row.fields.AccessDisplayName ?? email,
        role,
        org: row.fields.AccessOrg ?? undefined,
      });
    }
    cachedEmailRoleMap = map;
    cachedTeamMembers = members;
    cachedSource = 'sharepoint';
  } catch (err) {
    // eslint-disable-next-line no-console
    console.warn('[Access List] Failed to load from SharePoint, using hardcoded fallback:', err);
    cachedEmailRoleMap = { ...FALLBACK_EMAIL_ROLE_MAP };
    cachedTeamMembers = [...FALLBACK_TEAM_MEMBERS];
    cachedSource = 'fallback';
  }
}

/** Returns 'sharepoint' / 'fallback' / 'unloaded' — used by the Settings editor to show a banner. */
export function getAccessListSource(): 'sharepoint' | 'fallback' | 'unloaded' {
  return cachedSource;
}

/**
 * Synchronous role lookup. Reads from the cache populated by
 * `loadAccessList()`. Before the cache is warm, falls back to the
 * hardcoded map so a quick render attempt doesn't fail catastrophically.
 *
 * Returns null when the user is not on the list (= access denied).
 */
export function lookupRole(email: string | undefined | null): Role | null {
  if (!email) return null;
  const key = email.toLowerCase();
  const map = cachedEmailRoleMap ?? FALLBACK_EMAIL_ROLE_MAP;
  return map[key] ?? null;
}

/**
 * Compute display initials from a name (preferred) or email (fallback).
 *   "Brandy Turner" → "BT"
 *   "Bryan" → "BR"
 *   "lori@newshire.com" → "LO"
 */
export function getUserInitials(
  name: string | undefined | null,
  email: string | undefined | null,
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

/** Derive a display "org" label from email domain — cosmetic only. */
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

/**
 * Team member directory for assignee pickers + autocomplete.
 *
 * Backed by the same cache. Reading this before `loadAccessList()` resolves
 * returns the hardcoded fallback list — fine for tests and for the very
 * first render after sign-in.
 */
export const TEAM_MEMBERS: TeamMember[] = new Proxy([] as TeamMember[], {
  get(_target, prop) {
    const list = cachedTeamMembers ?? FALLBACK_TEAM_MEMBERS;
    // Reflect arrayish reads (length, indexed, iterator, map, forEach, etc.)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const value = (list as any)[prop];
    return typeof value === 'function' ? value.bind(list) : value;
  },
});
