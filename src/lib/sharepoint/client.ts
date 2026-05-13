import { Client } from '@microsoft/microsoft-graph-client';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { msalInstance } from '../auth/msalConfig';

/**
 * Scopes required for SharePoint access.
 *
 * These are admin-consented at the tenant level on the CAHP Compliance Hub Azure AD app,
 * so users receive tokens silently without a consent prompt.
 */
const SHAREPOINT_SCOPES = ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All'];

/**
 * Cached site ID — resolved once per session from VITE_SHAREPOINT_SITE.
 * The site ID is a stable GUID compound; resolving by path is a 1-hop lookup we
 * do once and cache for the duration of the session.
 */
let cachedSiteId: string | null = null;

/**
 * Acquire an access token for SharePoint scopes via MSAL.
 *
 * Silent first; falls back to redirect-based interaction if the user needs to consent
 * or re-authenticate. Throws if no account is signed in.
 */
async function getAccessToken(): Promise<string> {
  const account = msalInstance.getActiveAccount();
  if (!account) {
    throw new Error('Not signed in. Sign-in is handled by SignInGate before this code runs.');
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: SHAREPOINT_SCOPES,
      account,
    });
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      // User needs to consent or re-auth. Redirect flow.
      await msalInstance.acquireTokenRedirect({
        scopes: SHAREPOINT_SCOPES,
        account,
      });
      // acquireTokenRedirect navigates away; this throw is unreachable in practice
      throw new Error('Redirecting for consent…');
    }
    throw err;
  }
}

/**
 * Singleton Graph client. Used internally by the typed helpers below — components
 * should NOT call this directly. Use getListItems / getListItem / createListItem etc.
 */
const graphClient = Client.init({
  authProvider: (done) => {
    getAccessToken()
      .then((token) => done(null, token))
      .catch((err) => done(err as Error, null));
  },
});

/**
 * Resolve the SharePoint site ID from the VITE_SHAREPOINT_SITE env var.
 * Format: hostname:/path  (e.g., `vanrockre.sharepoint.com:/sites/CAHPComplianceHub`).
 */
export async function getSiteId(): Promise<string> {
  if (cachedSiteId) return cachedSiteId;

  const sitePath = import.meta.env.VITE_SHAREPOINT_SITE;
  if (!sitePath) {
    throw new Error(
      'VITE_SHAREPOINT_SITE is not configured. Add it to .env.local locally and as a GitHub Actions repository variable for production. Format: `<tenant>.sharepoint.com:/sites/<SiteName>`.'
    );
  }

  const site: { id: string } = await graphClient.api(`/sites/${sitePath}`).get();
  cachedSiteId = site.id;
  return cachedSiteId;
}

// =============================================================================
// LIST OPERATIONS
// =============================================================================

export interface ListQueryOptions {
  /** OData filter string, e.g. `fields/cahpState eq 'SC'` */
  filter?: string;
  /** OData orderby, e.g. `fields/Title asc` */
  orderBy?: string;
  /** Max items to return. SharePoint default is 100. */
  top?: number;
  /** Expand lookup fields (defaults to expanding `fields`) */
  expand?: string;
}

/**
 * Get all items from a SharePoint list, with `fields` expanded by default.
 * For large lists, pass `top` to limit results — the app should paginate elsewhere
 * for lists that grow past a few hundred items.
 */
export async function getListItems<TItem>(
  listName: string,
  options: ListQueryOptions = {}
): Promise<TItem[]> {
  const siteId = await getSiteId();
  let request = graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`)
    .expand(options.expand ?? 'fields');

  if (options.filter) request = request.filter(options.filter);
  if (options.orderBy) request = request.orderby(options.orderBy);
  if (options.top) request = request.top(options.top);

  const response: { value: TItem[] } = await request.get();
  return response.value;
}

/** Get a single list item by its SharePoint integer ID (passed as a string). */
export async function getListItem<TItem>(
  listName: string,
  itemId: string
): Promise<TItem> {
  const siteId = await getSiteId();
  return graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}`)
    .expand('fields')
    .get();
}

/**
 * Create a new list item. The `fields` parameter is the column values.
 * Lookup columns are set by `{FieldName}LookupId` with the target item's integer ID.
 * Returns the created item with its assigned ID.
 *
 * Auto-logs the CREATE to the AuditLog list. The AuditLog list itself is excluded
 * from auditing (no recursion).
 */
export async function createListItem<TItem>(
  listName: string,
  fields: Record<string, unknown>
): Promise<TItem> {
  const siteId = await getSiteId();
  const created = await graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`)
    .post({ fields });

  await tryAuditLog({
    listName,
    action: 'CREATE',
    itemId: (created as { id: string }).id,
    after: fields,
    entityTitle: (fields.Title as string) ?? '(no title)',
  });

  return created;
}

/**
 * Update an existing list item's column values.
 * Only included fields are modified; omitted fields are left untouched.
 *
 * Auto-logs the UPDATE to the AuditLog list with a field-by-field diff. The AuditLog
 * list itself is excluded from auditing.
 */
export async function updateListItem<TItem>(
  listName: string,
  itemId: string,
  fields: Record<string, unknown>
): Promise<TItem> {
  const siteId = await getSiteId();

  // Capture before-state for the audit diff
  let beforeFields: Record<string, unknown> | undefined;
  if (listName !== 'AuditLog') {
    try {
      const beforeItem = await getListItem<{ fields: Record<string, unknown> }>(listName, itemId);
      beforeFields = beforeItem.fields;
    } catch {
      // If pre-read fails, continue with the update anyway — audit just won't have before-state
    }
  }

  const updated = await graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`)
    .patch(fields);

  await tryAuditLog({
    listName,
    action: 'UPDATE',
    itemId,
    before: beforeFields,
    after: fields,
    entityTitle: (beforeFields?.Title as string) ?? (fields.Title as string) ?? '(no title)',
  });

  return updated;
}

/**
 * Delete a list item permanently.
 * Use sparingly — most "deletes" in the app should be soft-archives.
 *
 * Auto-logs the DELETE to the AuditLog list (with full pre-delete state for forensics).
 */
export async function deleteListItem(listName: string, itemId: string): Promise<void> {
  const siteId = await getSiteId();

  // Capture record before deletion for audit forensics
  let beforeFields: Record<string, unknown> | undefined;
  let entityTitle = '(no title)';
  if (listName !== 'AuditLog') {
    try {
      const beforeItem = await getListItem<{ fields: Record<string, unknown> }>(listName, itemId);
      beforeFields = beforeItem.fields;
      entityTitle = (beforeFields.Title as string) ?? entityTitle;
    } catch {
      // continue
    }
  }

  await graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}`)
    .delete();

  await tryAuditLog({
    listName,
    action: 'DELETE',
    itemId,
    before: beforeFields,
    entityTitle,
  });
}

// =============================================================================
// AUDIT LOGGING
// =============================================================================

/** Map SharePoint list names to human-readable entity type labels for the audit log. */
const ENTITY_TYPE_LABELS: Record<string, string> = {
  'Properties Registry': 'Property',
  'Submittals Tracker': 'Submittal',
  'Compliance Deadlines': 'Compliance Deadline',
  'DOR Correspondence Log': 'DOR Correspondence',
  'Billing Tracker': 'Billing Entry',
  'Outstanding Items Checklist': 'Outstanding Item',
  'Known Issues Log': 'Known Issue',
  'Ownership Structure': 'Ownership Record',
  'Property Notes': 'Note',
};

interface AuditLogInput {
  listName: string;
  action: 'CREATE' | 'UPDATE' | 'DELETE';
  itemId: string;
  before?: Record<string, unknown>;
  after?: Record<string, unknown>;
  entityTitle: string;
}

/**
 * Write one row to the AuditLog SharePoint list.
 * Failures are caught and logged to console — never bubble up because we never want
 * a failed audit write to prevent the user's data operation from succeeding.
 */
async function tryAuditLog(input: AuditLogInput): Promise<void> {
  // Don't audit the audit log itself (would recurse)
  if (input.listName === 'AuditLog') return;

  const entityType = ENTITY_TYPE_LABELS[input.listName] ?? input.listName;
  const summary = buildChangeSummary(input);
  const fields: Record<string, unknown> = {
    Title: `${input.action === 'CREATE' ? 'Created' : input.action === 'UPDATE' ? 'Updated' : 'Deleted'} ${entityType}: ${input.entityTitle}`,
    Action: input.action,
    EntityType: entityType,
    EntityId: input.itemId,
    EntityTitle: input.entityTitle,
    ChangeSummary: summary,
  };

  if (input.before) {
    fields.BeforeJSON = safeStringify(input.before);
  }
  if (input.after) {
    fields.AfterJSON = safeStringify(input.after);
  }

  try {
    const siteId = await getSiteId();
    await graphClient
      .api(`/sites/${siteId}/lists/AuditLog/items`)
      .post({ fields });
  } catch (err) {
    // Log to console but don't throw — the underlying data operation already succeeded
    // and we don't want to confuse the user with a "save failed" message they wouldn't understand.
    // eslint-disable-next-line no-console
    console.warn('[AuditLog] Failed to write audit entry:', err, input);
  }
}

/**
 * Build a human-readable summary of what changed.
 * For UPDATE: lists each changed field with old → new values.
 * For CREATE/DELETE: returns a brief description.
 */
function buildChangeSummary(input: AuditLogInput): string {
  if (input.action === 'CREATE') {
    return `New record created.`;
  }
  if (input.action === 'DELETE') {
    return `Record permanently deleted.`;
  }
  // UPDATE — compute field-by-field diff
  if (!input.before || !input.after) return '(no diff available)';

  const lines: string[] = [];
  const fieldsToCheck = Object.keys(input.after);
  for (const field of fieldsToCheck) {
    const oldVal = input.before[field];
    const newVal = input.after[field];
    if (oldVal !== newVal) {
      lines.push(`${field}: "${formatValue(oldVal)}" → "${formatValue(newVal)}"`);
    }
  }
  return lines.length > 0 ? lines.join('\n') : '(no changes detected)';
}

function formatValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' && v.length > 100) return v.slice(0, 100) + '…';
  return String(v);
}

function safeStringify(obj: unknown): string {
  try {
    const json = JSON.stringify(obj, null, 2);
    // Truncate if huge — SharePoint multi-line text has a practical limit
    if (json.length > 50000) return json.slice(0, 50000) + '\n…[truncated]';
    return json;
  } catch {
    return '(unserializable)';
  }
}
