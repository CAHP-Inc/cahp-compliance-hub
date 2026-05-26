import { Client } from '@microsoft/microsoft-graph-client';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { msalInstance } from '../auth/msalConfig';

/**
 * Scopes required for Graph access.
 *
 * These are admin-consented at the tenant level on the CAHP Compliance Hub Azure AD app,
 * so users receive tokens silently without a consent prompt.
 *
 * Mail.Send is needed by the in-app email composer (lib/email.ts). It only
 * lets the app send messages as the signed-in user — no read, no folder
 * access, no impersonation.
 */
const SHAREPOINT_SCOPES = ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All', 'Mail.Send'];

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
 * Singleton Graph client. Used internally by the typed helpers below.
 *
 * Most components should use the typed helpers (getListItems, createListItem,
 * etc.). The `graphClient` is exported only so adjacent Graph-backed
 * features — like lib/email.ts's sendMail() — can reuse the same auth +
 * token-refresh plumbing.
 */
export const graphClient = Client.init({
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
  fields: Record<string, unknown>,
  options?: { reason?: string }
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
    reason: options?.reason,
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
// DOCUMENT LIBRARY OPERATIONS (PR-11c)
// =============================================================================
//
// SharePoint document libraries are exposed via Graph as both Lists and Drives:
//   - List view: `/sites/{siteId}/lists/{libraryName}/items` — gives metadata items
//   - Drive view: `/sites/{siteId}/drives/{driveId}/...` — gives file operations
//
// To upload a file: PUT to /sites/{siteId}/drives/{driveId}/root:/{filename}:/content
// Then PATCH the corresponding listItem/fields to set PropertyLookupId metadata.
// =============================================================================

/** Cached drive IDs by library name — drives are stable per library, lookup once. */
const driveIdCache = new Map<string, string>();

async function getDriveIdForLibrary(libraryName: string): Promise<string> {
  if (driveIdCache.has(libraryName)) return driveIdCache.get(libraryName)!;
  const siteId = await getSiteId();
  const response: { drive: { id: string } } = await graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(libraryName)}`)
    .expand('drive')
    .get();
  if (!response.drive?.id) {
    throw new Error(`Library "${libraryName}" has no associated drive`);
  }
  driveIdCache.set(libraryName, response.drive.id);
  return response.drive.id;
}

export interface DocumentUploadOptions {
  libraryName: string;
  filename: string;
  file: File | Blob;
  /** Metadata fields to set on the uploaded document's listItem (PropertyLookupId, etc.) */
  metadata?: Record<string, unknown>;
  /** Progress callback (0-100) — currently fires only at start (0) and end (100); chunked upload TBD */
  onProgress?: (percent: number) => void;
}

export interface DocumentUploadResult {
  itemId: string;
  driveItemId: string;
  webUrl: string;
  filename: string;
  size: number;
}

/**
 * Upload a file to a SharePoint Document Library and set metadata.
 *
 * Uses Graph's two upload paths automatically based on file size:
 *   - Files ≤ 4 MB: single PUT (simple upload, 1 round-trip)
 *   - Files > 4 MB: upload session with 10 MiB chunks
 *
 * Chunked uploads report real per-chunk progress through `onProgress`.
 * On chunk failure the upload aborts — caller can retry by invoking again.
 */
export async function uploadDocument(options: DocumentUploadOptions): Promise<DocumentUploadResult> {
  const { libraryName, filename, file, metadata, onProgress } = options;

  onProgress?.(0);

  const driveId = await getDriveIdForLibrary(libraryName);
  const encodedName = encodeURIComponent(filename);

  // Pick upload path by size
  const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024; // 4 MB
  let uploaded: { id: string; webUrl: string; size: number; name: string };

  if (file.size <= SIMPLE_UPLOAD_LIMIT) {
    // ─── Simple upload — single PUT ──────────────────────────
    uploaded = await graphClient
      .api(`/drives/${driveId}/root:/${encodedName}:/content`)
      .header('Content-Type', file.type || 'application/octet-stream')
      .put(file);
    onProgress?.(100);
  } else {
    // ─── Chunked upload — upload session ──────────────────────
    uploaded = await uploadInChunks(driveId, encodedName, file, onProgress);
  }

  // Set metadata if provided. Different libraries have inconsistent lookup
  // column internal names ('OwnerLookupId', 'Owner', 'OwnerId' — same for
  // Property), so we try each variant and verify the value actually persisted
  // by re-reading. Without this, a PATCH against a non-existent column quietly
  // succeeds, the file lands tagged-as-nothing, and EntityDocumentsSection's
  // filter hides it — exactly the 'upload disappears into the void' symptom.
  let listItemId: string | null = null;
  if (metadata && Object.keys(metadata).length > 0) {
    try {
      listItemId = await patchListItemFieldsWithFallback(
        driveId,
        uploaded.id,
        metadata,
        libraryName,
      );
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      // eslint-disable-next-line no-console
      console.warn('Metadata patch failed after upload:', e);
      throw new Error(
        `File uploaded to "${libraryName}" but its Owner/Property tag couldn't be saved — without the tag the doc won't appear on the entity's Documents tab. ` +
        `Run scripts/provision-document-library-columns.ps1 to add the missing lookup columns to your libraries, then re-upload. (Cause: ${msg})`,
      );
    }
  }

  // If we don't have a listItemId from the patch, fetch it
  if (!listItemId) {
    try {
      const li: { id: string } = await graphClient
        .api(`/drives/${driveId}/items/${uploaded.id}/listItem`)
        .get();
      listItemId = li.id;
    } catch {
      listItemId = uploaded.id; // fallback
    }
  }

  // Audit log
  await tryAuditLog({
    listName: libraryName,
    action: 'CREATE',
    itemId: listItemId,
    after: { Title: uploaded.name, ...(metadata ?? {}) },
    entityTitle: uploaded.name,
  });

  return {
    itemId: listItemId,
    driveItemId: uploaded.id,
    webUrl: uploaded.webUrl,
    filename: uploaded.name,
    size: uploaded.size,
  };
}

/**
 * PATCH a freshly-uploaded driveItem's listItem fields, trying alternate
 * internal names for the lookup columns. SharePoint document libraries have
 * been provisioned inconsistently across the deployment, so the same logical
 * tag ('owner = 42') might map to a column named 'Owner', 'OwnerLookupId',
 * or 'OwnerId' depending on the library. We try each, then verify the value
 * actually persisted by reading the item back.
 */
async function patchListItemFieldsWithFallback(
  driveId: string,
  driveItemId: string,
  metadata: Record<string, unknown>,
  libraryName: string,
): Promise<string | null> {
  // Split metadata into "needs-fallback" (lookup IDs) and "static" (everything else)
  const lookupAliases: Record<string, string[]> = {
    OwnerLookupId: ['OwnerLookupId', 'Owner', 'OwnerId'],
    PropertyLookupId: ['PropertyLookupId', 'Property', 'PropertyId'],
  };

  const staticFields: Record<string, unknown> = {};
  const lookupFields: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(metadata)) {
    if (key in lookupAliases) lookupFields[key] = value;
    else staticFields[key] = value;
  }

  let listItemId: string | null = null;

  // Helper — does a single PATCH-then-readback cycle, returning whether the
  // value persisted. Tries both single-value and multi-value (array) shapes
  // since the column might be either.
  //
  // Multi-value Lookup readback gotcha: Graph returns the field under its
  // base name (e.g. "Owner") as an array of {LookupId, LookupValue} objects,
  // NOT under "OwnerLookupId". So we check several candidate readback keys
  // and inspect both scalars and object-arrays.
  const tryWriteAndVerify = async (
    fieldName: string,
    rawValue: unknown,
    asArray: boolean,
    candidateReadbackKeys: string[],
  ): Promise<{ persisted: boolean; readBack: unknown }> => {
    const payload: Record<string, unknown> = asArray
      ? {
          [`${fieldName}@odata.type`]: 'Collection(Edm.Int32)',
          [fieldName]: [Number(rawValue)],
        }
      : { [fieldName]: rawValue };
    try {
      const updated: { id: string } = await graphClient
        .api(`/drives/${driveId}/items/${driveItemId}/listItem/fields`)
        .patch(payload);
      listItemId = updated.id ?? listItemId;
    } catch (e) {
      // eslint-disable-next-line no-console
      console.warn(`[upload] ${libraryName}: PATCH "${fieldName}" (${asArray ? 'array' : 'single'}) threw:`, e);
      return { persisted: false, readBack: undefined };
    }

    // Read back
    const verify: { fields: Record<string, unknown> } = await graphClient
      .api(`/drives/${driveId}/items/${driveItemId}/listItem`)
      .get();
    const expected = String(rawValue ?? '');

    // Check every candidate readback key; first match wins. Each value can be:
    //   - a scalar (single-value LookupId)
    //   - an array of scalars (some multi-value shapes)
    //   - an array of {LookupId, LookupValue} (the common multi-value shape)
    let readBack: unknown = undefined;
    for (const key of candidateReadbackKeys) {
      const v = verify.fields[key];
      if (v === undefined || v === null || v === '') continue;
      readBack = v;
      if (Array.isArray(v)) {
        for (const entry of v) {
          if (entry && typeof entry === 'object' && 'LookupId' in entry) {
            if (String((entry as { LookupId: unknown }).LookupId) === expected) {
              return { persisted: true, readBack: v };
            }
          } else if (String(entry) === expected) {
            return { persisted: true, readBack: v };
          }
        }
      } else if (String(v) === expected) {
        return { persisted: true, readBack: v };
      }
    }
    return { persisted: false, readBack };
  };

  // Try each lookup field. For each, attempt every alias × {single, array}
  // shape until one combo persists.
  for (const [primaryKey, value] of Object.entries(lookupFields)) {
    const aliases = lookupAliases[primaryKey];
    const expected = String(value ?? '');
    // Candidate keys to look at on readback: every alias PLUS the base column
    // name (without the LookupId suffix). Multi-value Lookups surface under
    // the base name, e.g. PATCH "OwnerLookupId" but read "Owner".
    const baseName = primaryKey.replace(/LookupId$/, '');
    const readbackKeys = Array.from(new Set([...aliases, baseName]));
    let succeededWith: string | null = null;

    outer: for (const alias of aliases) {
      for (const asArray of [false, true]) {
        const { persisted, readBack } = await tryWriteAndVerify(alias, value, asArray, readbackKeys);
        if (persisted) {
          succeededWith = `${alias}${asArray ? '[]' : ''}`;
          // eslint-disable-next-line no-console
          console.log(`[upload] ${libraryName}: tagged ${primaryKey}=${expected} via field "${succeededWith}"`);
          break outer;
        }
        // eslint-disable-next-line no-console
        console.warn(`[upload] ${libraryName}: PATCH "${alias}" (${asArray ? 'array' : 'single'}) returned OK but value did not persist (read back: ${JSON.stringify(readBack)})`);
      }
    }

    if (!succeededWith) {
      throw new Error(
        `Tried [${aliases.join(', ')}] in both single and array form on library "${libraryName}" — none persisted the ${primaryKey} value. The column may be missing or read-only.`,
      );
    }
  }

  // Apply any remaining non-lookup static fields in one PATCH
  if (Object.keys(staticFields).length > 0) {
    const updated: { id: string } = await graphClient
      .api(`/drives/${driveId}/items/${driveItemId}/listItem/fields`)
      .patch(staticFields);
    listItemId = updated.id ?? listItemId;
  }

  return listItemId;
}

/**
 * Chunked upload via Graph upload session.
 *
 * Per Microsoft docs: chunk sizes must be multiples of 320 KiB (327,680 bytes).
 * 10 MiB is the sweet spot for most networks — large enough to amortize HTTP overhead
 * but small enough that one failed chunk doesn't waste much progress.
 *
 * The upload session URL is pre-signed by Graph at session creation time, so chunk
 * PUTs do NOT require an auth header — they use the bare `fetch` API rather than graphClient.
 */
async function uploadInChunks(
  driveId: string,
  encodedName: string,
  file: File | Blob,
  onProgress?: (percent: number) => void
): Promise<{ id: string; webUrl: string; size: number; name: string }> {
  const CHUNK_SIZE = 10 * 1024 * 1024; // 10 MiB — multiple of 320 KiB

  // 1. Create upload session — returns a pre-signed uploadUrl
  const session: { uploadUrl: string; expirationDateTime: string } = await graphClient
    .api(`/drives/${driveId}/root:/${encodedName}:/createUploadSession`)
    .post({
      item: {
        '@microsoft.graph.conflictBehavior': 'rename', // auto-rename on collision
      },
    });

  const uploadUrl = session.uploadUrl;
  let finalItem: { id: string; webUrl: string; size: number; name: string } | null = null;

  // 2. PUT each chunk with Content-Range header
  for (let offset = 0; offset < file.size; offset += CHUNK_SIZE) {
    const chunkEnd = Math.min(offset + CHUNK_SIZE, file.size);
    const chunk = file.slice(offset, chunkEnd);
    const contentRange = `bytes ${offset}-${chunkEnd - 1}/${file.size}`;

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: { 'Content-Range': contentRange },
      body: chunk,
    });

    if (!response.ok) {
      // Best-effort cleanup — DELETE the session so we don't leak partial uploads
      try {
        await fetch(uploadUrl, { method: 'DELETE' });
      } catch {
        // ignore — session will expire on its own
      }
      throw new Error(
        `Chunk upload failed at byte ${offset} (status ${response.status}). ` +
          `Retry by clicking Upload again.`
      );
    }

    // Progress: percent of bytes uploaded
    onProgress?.(Math.round((chunkEnd / file.size) * 100));

    // The final chunk's response body contains the completed driveItem
    if (chunkEnd === file.size) {
      finalItem = (await response.json()) as typeof finalItem;
    }
    // Intermediate chunks return 202 Accepted with `nextExpectedRanges` — we don't need it
    // because we're uploading sequentially without parallelism. Body is discarded.
  }

  if (!finalItem) {
    throw new Error('Upload finished but Graph did not return a driveItem.');
  }
  return finalItem;
}



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
  'Owners': 'Owner',
  'Disbursements': 'Disbursement',
  'Owner Communications': 'Communication',
  'Notifications': 'Notification',
};

interface AuditLogInput {
  listName: string;
  action: 'CREATE' | 'UPDATE' | 'DELETE';
  itemId: string;
  before?: Record<string, unknown>;
  after?: Record<string, unknown>;
  entityTitle: string;
  reason?: string;          // User-supplied reason (e.g., ownership change: Buy-In, Buy-Out, Estate, etc.)
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
  if (!input.before || !input.after) {
    return input.reason ? `(no diff available) — Reason: ${input.reason}` : '(no diff available)';
  }

  const lines: string[] = [];
  const fieldsToCheck = Object.keys(input.after);
  for (const field of fieldsToCheck) {
    const oldVal = input.before[field];
    const newVal = input.after[field];
    if (oldVal !== newVal) {
      lines.push(`${field}: "${formatValue(oldVal)}" → "${formatValue(newVal)}"`);
    }
  }
  const summary = lines.length > 0 ? lines.join('\n') : '(no changes detected)';
  return input.reason ? `${summary}\n— Reason: ${input.reason}` : summary;
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
