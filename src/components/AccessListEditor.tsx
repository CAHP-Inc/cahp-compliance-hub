import { useEffect, useMemo, useState } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type AccessListEntry,
  type AccessRole,
} from '../lib/sharepoint';
import { useSession } from '../lib/session';
import { getAccessListSource } from '../lib/roleMap';
import { Icon } from './ui/Icon';

/**
 * Settings → Access List editor.
 *
 * Replaces the previous read-only display backed by a hardcoded
 * EMAIL_ROLE_MAP. Admins add/edit/remove users + roles here without
 * needing to edit code or redeploy.
 *
 * On save, the in-memory role cache (lib/roleMap.ts) is refreshed via
 * session.refreshAccess() so the change takes effect immediately for
 * the editing user. Other signed-in users pick up the change on their
 * next page refresh (the access list is loaded once per session).
 */

const ROLES: AccessRole[] = ['Admin', 'Contributor', 'Accounting'];

const ROLE_STYLES: Record<AccessRole, string> = {
  Admin: 'bg-gold-500 text-teal-900 border-gold-700',
  Contributor: 'bg-teal-700 text-white border-teal-900',
  Accounting: 'bg-blue-600 text-white border-blue-800',
};

interface EditorRow {
  rowId: string | null;
  email: string;
  role: AccessRole;
  displayName: string;
  org: string;
  active: boolean;
  notes: string;
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function AccessListEditor() {
  const session = useSession();
  const list = useSharePointList<AccessListEntry>(LIST_NAMES.AccessList, { top: 500 });
  const [items, setItems] = useState<EditorRow[]>([]);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [savedStamp, setSavedStamp] = useState<string | null>(null);

  // Hydrate working copy
  useEffect(() => {
    if (list.loading) return;
    if (list.data) {
      const sorted = [...list.data].sort((a, b) =>
        (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
      );
      setItems(
        sorted.map((row) => ({
          rowId: row.id,
          email: row.fields.Title ?? '',
          role: row.fields.AccessRole ?? 'Contributor',
          displayName: row.fields.AccessDisplayName ?? '',
          org: row.fields.AccessOrg ?? '',
          active: row.fields.AccessActive !== false,
          notes: row.fields.AccessNotes ?? '',
        })),
      );
    }
  }, [list.loading, list.data]);

  const source = useMemo(() => getAccessListSource(), [savedStamp]);

  const updateItem = (idx: number, patch: Partial<EditorRow>) => {
    setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, ...patch } : it)));
  };
  const removeItem = (idx: number) => setItems((prev) => prev.filter((_, i) => i !== idx));
  const addItem = () =>
    setItems((prev) => [
      ...prev,
      { rowId: null, email: '', role: 'Contributor', displayName: '', org: '', active: true, notes: '' },
    ]);

  // Validation
  const validationErrors = useMemo(() => {
    const errs: string[] = [];
    const seen = new Set<string>();
    for (const row of items) {
      const email = row.email.trim().toLowerCase();
      if (!email) continue;
      if (!EMAIL_RE.test(email)) errs.push(`Invalid email: ${row.email}`);
      if (seen.has(email)) errs.push(`Duplicate email: ${row.email}`);
      seen.add(email);
    }
    // Safety: don't let the current user lock themselves out
    const myEmail = session.user?.email?.toLowerCase();
    if (myEmail) {
      const me = items.find((r) => r.email.trim().toLowerCase() === myEmail);
      if (me) {
        if (!me.active) errs.push("You're about to deactivate yourself — change someone else's role first if you need to step down.");
        if (me.role !== 'Admin') errs.push("You're about to remove your own Admin role — make sure another Admin exists first.");
      } else {
        errs.push("Don't remove yourself from the list — you'd lose Admin access immediately on save.");
      }
    }
    return errs;
  }, [items, session.user?.email]);

  const dirty = useMemo(() => {
    if (list.loading) return false;
    const existingMap = new Map<string, AccessListEntry>();
    (list.data ?? []).forEach((r) => existingMap.set(r.id, r));
    if (items.length !== (list.data?.length ?? 0)) return true;
    for (const row of items) {
      if (!row.rowId) return true;
      const sp = existingMap.get(row.rowId);
      if (!sp) return true;
      const f = sp.fields;
      if (
        (f.Title ?? '') !== row.email ||
        f.AccessRole !== row.role ||
        (f.AccessDisplayName ?? '') !== row.displayName ||
        (f.AccessOrg ?? '') !== row.org ||
        (f.AccessActive !== false) !== row.active ||
        (f.AccessNotes ?? '') !== row.notes
      ) {
        return true;
      }
    }
    return false;
  }, [items, list.data, list.loading]);

  const handleSave = async () => {
    if (validationErrors.length > 0) return;
    const cleaned = items.filter((i) => i.email.trim().length > 0);
    setSaving(true);
    setSaveError(null);
    try {
      const existingById = new Map<string, AccessListEntry>();
      (list.data ?? []).forEach((r) => existingById.set(r.id, r));
      const keptIds = new Set<string>();

      for (const row of cleaned) {
        const fields = {
          Title: row.email.trim().toLowerCase(),
          AccessRole: row.role,
          AccessDisplayName: row.displayName.trim() || null,
          AccessOrg: row.org.trim() || null,
          AccessActive: row.active,
          AccessNotes: row.notes.trim() || null,
        };
        if (row.rowId && existingById.has(row.rowId)) {
          keptIds.add(row.rowId);
          await updateListItem(LIST_NAMES.AccessList, row.rowId, fields);
        } else {
          await createListItem(LIST_NAMES.AccessList, fields);
        }
      }
      for (const [id] of existingById) {
        if (!keptIds.has(id)) await deleteListItem(LIST_NAMES.AccessList, id);
      }

      list.refetch?.();
      // Refresh the in-memory role cache so the change takes effect now
      await session.refreshAccess();
      setSavedStamp(new Date().toISOString());
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-4">
      {list.error ? (
        <div className="bg-red-50 border border-red-200 rounded-md p-3 text-xs text-red-900">
          <strong>Couldn't reach the Access List in SharePoint.</strong>
          <div className="mt-1">{list.error.message}</div>
          <div className="mt-2">
            Run <code className="bg-white px-1 rounded">scripts\provision-access-list.ps1</code>{' '}
            to create the list. While the list is unreachable the app uses a hardcoded fallback so the original team isn't locked out.
          </div>
        </div>
      ) : source === 'fallback' ? (
        <div className="bg-amber-50 border border-amber-200 rounded-md p-3 text-xs text-amber-900">
          <strong>Using hardcoded fallback.</strong> The SharePoint Access List is empty or hasn't been provisioned. Edits saved here will become the source of truth.
        </div>
      ) : (
        <div className="bg-emerald-50 border border-emerald-200 rounded-md p-3 text-xs text-emerald-900">
          <strong>Synced to SharePoint.</strong> New entries grant access immediately. Existing signed-in users pick up role changes on their next page refresh.
        </div>
      )}

      {validationErrors.length > 0 && (
        <div className="bg-red-50 border border-red-200 rounded-md p-3 text-xs text-red-900 space-y-0.5">
          {validationErrors.map((e, i) => <div key={i}>• {e}</div>)}
        </div>
      )}

      <div className="bg-white border border-gray-200 rounded-lg shadow-card">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between gap-2 flex-wrap">
          <div>
            <h3 className="text-sm font-semibold text-teal-900">Access List</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {items.length} entr{items.length === 1 ? 'y' : 'ies'}
              {dirty && <span className="ml-2 text-amber-700 font-semibold">· unsaved changes</span>}
            </p>
          </div>
          <button
            onClick={handleSave}
            disabled={!dirty || saving || validationErrors.length > 0}
            className="text-xs px-3 py-1.5 rounded-md font-medium bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white inline-flex items-center gap-1.5"
            title={validationErrors.length > 0 ? 'Fix the validation errors above first' : ''}
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Saving…' : 'Save changes'}
          </button>
        </div>

        {saveError && (
          <div className="px-4 py-2 bg-red-50 border-b border-red-200 text-xs text-red-800">
            {saveError}
          </div>
        )}

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-[11px] font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-3 py-2 text-left">Email (UPN)</th>
                <th className="px-3 py-2 text-left w-44">Display Name</th>
                <th className="px-3 py-2 text-left w-32">Org</th>
                <th className="px-3 py-2 text-left w-36">Role</th>
                <th className="px-3 py-2 text-left w-20">Active</th>
                <th className="px-3 py-2 text-right w-20">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {items.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-3 py-6 text-center text-xs text-gray-500 italic">
                    No entries. Add one below.
                  </td>
                </tr>
              ) : (
                items.map((item, idx) => {
                  const isMe = session.user?.email?.toLowerCase() === item.email.trim().toLowerCase();
                  return (
                    <tr key={item.rowId ?? `new-${idx}`} className="hover:bg-gray-50">
                      <td className="px-3 py-2 align-top">
                        <input
                          type="email"
                          value={item.email}
                          onChange={(e) => updateItem(idx, { email: e.target.value })}
                          disabled={saving}
                          placeholder="user@domain.com"
                          className="w-full px-2 py-1 border border-gray-300 rounded text-xs font-mono-data focus:outline-none focus:border-teal-500"
                        />
                        {isMe && (
                          <span className="text-[10px] text-teal-700 font-semibold mt-1 inline-block">↑ you</span>
                        )}
                        <textarea
                          value={item.notes}
                          onChange={(e) => updateItem(idx, { notes: e.target.value })}
                          disabled={saving}
                          rows={1}
                          placeholder="Notes (optional)"
                          className="w-full mt-1 px-2 py-1 border border-gray-200 rounded text-[11px] text-gray-600 resize-none focus:outline-none focus:border-teal-500"
                        />
                      </td>
                      <td className="px-3 py-2 align-top">
                        <input
                          type="text"
                          value={item.displayName}
                          onChange={(e) => updateItem(idx, { displayName: e.target.value })}
                          disabled={saving}
                          placeholder="Friendly name"
                          className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500"
                        />
                      </td>
                      <td className="px-3 py-2 align-top">
                        <input
                          type="text"
                          value={item.org}
                          onChange={(e) => updateItem(idx, { org: e.target.value })}
                          disabled={saving}
                          placeholder="Org"
                          className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500"
                        />
                      </td>
                      <td className="px-3 py-2 align-top">
                        <select
                          value={item.role}
                          onChange={(e) => updateItem(idx, { role: e.target.value as AccessRole })}
                          disabled={saving}
                          className={`w-full px-2 py-1 border rounded text-xs font-semibold focus:outline-none focus:border-teal-500 ${ROLE_STYLES[item.role]}`}
                        >
                          {ROLES.map((r) => (
                            <option key={r} value={r}>{r}</option>
                          ))}
                        </select>
                      </td>
                      <td className="px-3 py-2 align-top">
                        <label className="inline-flex items-center gap-1 text-xs cursor-pointer">
                          <input
                            type="checkbox"
                            checked={item.active}
                            onChange={(e) => updateItem(idx, { active: e.target.checked })}
                            disabled={saving}
                          />
                          <span>{item.active ? 'Yes' : 'No'}</span>
                        </label>
                      </td>
                      <td className="px-3 py-2 align-top text-right">
                        <button
                          onClick={() => removeItem(idx)}
                          disabled={saving}
                          className="text-[11px] text-error hover:text-red-700 font-medium px-2 py-1 rounded hover:bg-red-50 disabled:opacity-50"
                        >
                          Remove
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        <div className="px-4 py-3 border-t border-gray-100 bg-gray-50 rounded-b-lg">
          <button
            onClick={addItem}
            disabled={saving}
            className="text-xs px-3 py-1.5 rounded-md border border-dashed border-teal-400 text-teal-700 hover:bg-teal-50 inline-flex items-center gap-1.5 disabled:opacity-50"
          >
            <Icon name="plus" size={12} />
            Add user
          </button>
        </div>
      </div>

      <div className="text-[11px] text-gray-500 leading-relaxed space-y-1">
        <p>
          <strong>Adding a user:</strong> enter their full M365 email (UPN), pick a role, set Active = Yes, save. They can sign in immediately —
          their first sign-in fetches this list and grants access on the spot.
        </p>
        <p>
          <strong>Removing access:</strong> uncheck Active. The row stays for audit history. Use Remove only when you want to delete the record entirely.
        </p>
        <p>
          <strong>Role descriptions:</strong> Admin = full access, Contributor = create + edit but not delete/approve,
          Accounting = read-mostly with full Billing access.
        </p>
      </div>
    </div>
  );
}
