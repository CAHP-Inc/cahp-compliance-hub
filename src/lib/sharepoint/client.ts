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
 */
export async function createListItem<TItem>(
  listName: string,
  fields: Record<string, unknown>
): Promise<TItem> {
  const siteId = await getSiteId();
  return graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`)
    .post({ fields });
}

/**
 * Update an existing list item's column values.
 * Only included fields are modified; omitted fields are left untouched.
 */
export async function updateListItem<TItem>(
  listName: string,
  itemId: string,
  fields: Record<string, unknown>
): Promise<TItem> {
  const siteId = await getSiteId();
  return graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`)
    .patch(fields);
}

/**
 * Delete a list item permanently.
 * Use sparingly — most "deletes" in the app should be soft-archives.
 */
export async function deleteListItem(listName: string, itemId: string): Promise<void> {
  const siteId = await getSiteId();
  await graphClient
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}`)
    .delete();
}
