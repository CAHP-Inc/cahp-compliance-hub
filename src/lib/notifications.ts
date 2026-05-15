/**
 * Notifications helper — PR-15a
 *
 * Single entry point for creating in-app notifications.
 *
 * Notifications are stored in the SharePoint "Notifications" list and
 * fetched by the Notification Bell + Notifications Page filtered by UPN.
 *
 * Use notifyUser() in cascade workflows when an action affects another user.
 * Failures are swallowed silently — notification creation should never block
 * the primary action.
 */

import { createListItem, LIST_NAMES } from './sharepoint';
import type { NotificationType } from './sharepoint';

export interface NotifyUserParams {
  /** Recipient UPN (email). If undefined/empty, no-op. */
  upn: string | undefined;
  /** Notification headline shown in the feed */
  title: string;
  type: NotificationType;
  /** Entity type referenced — for category icon */
  targetType?: string;
  /** Entity ID — used to build the click-through URL */
  targetId?: string;
  /** Relative URL (e.g. "/properties/123" or "/submittals/45") */
  url?: string;
}

/**
 * Create a notification for a single user. Fire-and-forget; errors logged but not thrown.
 *
 * Self-notifications (where upn matches the currently-authenticated user) are skipped
 * automatically by upstream callers; this helper doesn't try to determine "self" —
 * callers must filter.
 */
export async function notifyUser({
  upn,
  title,
  type,
  targetType,
  targetId,
  url,
}: NotifyUserParams): Promise<void> {
  if (!upn) return;
  const normalized = upn.trim().toLowerCase();
  if (!normalized || !normalized.includes('@')) return;

  try {
    const fields: Record<string, unknown> = {
      Title: title,
      NotifAssignedTo: normalized,
      NotifType: type,
      NotifIsRead: false,
    };
    if (targetType) fields.NotifTargetType = targetType;
    if (targetId) fields.NotifTargetId = targetId;
    if (url) fields.NotifUrl = url;

    await createListItem(LIST_NAMES.Notifications, fields);
  } catch (err) {
    // Don't surface — notification failure shouldn't block the action that triggered it
    // eslint-disable-next-line no-console
    console.warn('Notification creation failed:', err);
  }
}

/**
 * Resolve a free-text "AssignedTo" value to a UPN where possible.
 *
 * Outstanding Items store AssignedTo as free text — could be "Lori", "lori@...",
 * "Lori Heckman", etc. This function tries to map common variants to known UPNs.
 *
 * Returns the input as-is if it's already an email, or undefined if it can't be resolved.
 * The roleMap is the canonical list of known users.
 */
const KNOWN_USERS: { upn: string; firstName: string; lastName?: string }[] = [
  { upn: 'bturner@newshirepm.com', firstName: 'brandy', lastName: 'turner' },
  { upn: 'stan@vanrockre.com', firstName: 'stan' },
  { upn: 'bdebruin@redcedarhomes.com', firstName: 'bryan', lastName: 'debruin' },
  { upn: 'lheckman@redcedarhomes.com', firstName: 'lori', lastName: 'heckman' },
];

export function resolveAssigneeToUpn(assignedTo: string | undefined): string | undefined {
  if (!assignedTo) return undefined;
  const trimmed = assignedTo.trim().toLowerCase();
  if (!trimmed) return undefined;
  if (trimmed.includes('@')) return trimmed;

  // Try matching first name or "first last" against known users
  for (const u of KNOWN_USERS) {
    if (trimmed === u.firstName) return u.upn;
    if (u.lastName && trimmed === `${u.firstName} ${u.lastName}`) return u.upn;
    if (u.lastName && trimmed === u.lastName) return u.upn;
  }
  return undefined;
}
