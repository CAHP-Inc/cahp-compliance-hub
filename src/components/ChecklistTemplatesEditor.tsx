import { useMemo, useState, useRef, useEffect } from 'react';
import {
  DOR_FILING_CHECKLIST,
  useChecklistTemplates,
  itemToTemplateFields,
  readLocalChecklistOverride,
  clearLocalChecklistOverride,
  type FilingChecklistItem,
  type FilingChecklistScope,
} from '../lib/filing-checklist';
import {
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type ChecklistTemplate,
} from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES } from './UploadDocumentModal';
import { Icon } from './ui/Icon';
import type { ItemCategory, CahpState } from '../lib/sharepoint';

/**
 * Settings → Checklist Templates editor.
 *
 * Reads from + writes to the shared SharePoint `Checklist Templates` list, so
 * every teammate sees the same configuration. Loads the hardcoded defaults
 * if the list hasn't been provisioned yet or is empty.
 *
 * On Save, the editor diffs the working copy against the SharePoint rows
 * (add new, update changed, delete removed) so partial failures don't leave
 * inconsistent state. Sort order is rewritten on every save.
 */

const ITEM_CATEGORIES: ItemCategory[] = [
  'Operating Agreement',
  'Articles of Incorporation',
  'EIN Confirmation',
  'Certificate of Existence',
  'Certificate of Authorization',
  '501(c)(3) Determination',
  'Deed',
  'Rent Roll',
  'LURA',
  'AMI Certification',
  'Org Chart',
  'Income Documentation',
  'Signed Submittal',
  'Determination Letter',
  'Other',
];

const SCOPES: { value: FilingChecklistScope; label: string; help: string }[] = [
  { value: 'cahp',     label: 'CAHP entity',      help: 'Doc lives on the CAHP nonprofit/LLC — reused across every filing' },
  { value: 'owner',    label: 'Property owner',   help: 'Doc lives on the property-owning LLC — reused per filing' },
  { value: 'property', label: 'Property',         help: 'Doc lives on the specific property — needed each filing' },
];

const LIBRARY_OPTIONS: { value: string; label: string }[] = [
  { value: '', label: 'Use category default' },
  ...PROPERTY_LINKED_LIBRARIES.map((l) => ({ value: l, label: l })),
];

const STATE_OPTIONS: { value: '' | CahpState; label: string }[] = [
  { value: '',   label: 'All states' },
  { value: 'SC', label: 'SC only' },
  { value: 'NC', label: 'NC only' },
];

/**
 * Editor row — the working copy. `rowId` ties back to a SharePoint listItem
 * when one exists; `null` means it's a newly-added row not yet persisted.
 */
interface EditorRow extends FilingChecklistItem {
  rowId: string | null;
}

export function ChecklistTemplatesEditor() {
  const { templates, rawRows, loading, error, usingFallback, refetch } = useChecklistTemplates();

  const [items, setItems] = useState<EditorRow[]>([]);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [savedStamp, setSavedStamp] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [showImportConfirm, setShowImportConfirm] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Pull live data into the local working copy whenever the SP list refreshes.
  // We index editor rows by the SharePoint listItem ID; newly added rows have
  // rowId === null and only get an ID once saved.
  useEffect(() => {
    if (loading) return;
    if (rawRows && rawRows.length > 0) {
      const sorted = [...rawRows].sort((a, b) => {
        const aOrder = a.fields.TemplateSortOrder ?? Number.MAX_SAFE_INTEGER;
        const bOrder = b.fields.TemplateSortOrder ?? Number.MAX_SAFE_INTEGER;
        return aOrder - bOrder;
      });
      setItems(
        sorted.map((row) => ({
          rowId: row.id,
          title: row.fields.Title ?? '',
          category: (row.fields.TemplateCategory as ItemCategory) ?? 'Other',
          scope: row.fields.TemplateScope ?? 'property',
          notes: row.fields.TemplateNotes,
          library: row.fields.TemplateLibrary,
          state: row.fields.TemplateState,
        })),
      );
    } else if (usingFallback) {
      // Seed the working copy with the hardcoded defaults so the editor isn't
      // blank when no SP rows exist yet. Saving these promotes them to SP.
      setItems(templates.map((t) => ({ ...t, rowId: null })));
    }
  }, [loading, rawRows, usingFallback, templates]);

  // Detect a local override that hasn't been imported into SharePoint yet
  const localOverride = useMemo(() => readLocalChecklistOverride(), [savedStamp]);

  const updateItem = (idx: number, patch: Partial<EditorRow>) => {
    setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, ...patch } : it)));
  };

  const removeItem = (idx: number) => {
    setItems((prev) => prev.filter((_, i) => i !== idx));
  };

  const addItem = () => {
    setItems((prev) => [
      ...prev,
      { rowId: null, title: '', category: 'Other', scope: 'property', notes: '' },
    ]);
  };

  const moveItem = (idx: number, dir: -1 | 1) => {
    setItems((prev) => {
      const next = [...prev];
      const j = idx + dir;
      if (j < 0 || j >= next.length) return prev;
      [next[idx], next[j]] = [next[j], next[idx]];
      return next;
    });
  };

  const handleSave = async () => {
    const cleaned = items.filter((i) => i.title.trim().length > 0);
    setSaving(true);
    setSaveError(null);
    try {
      // Snapshot of current SP rows for diff-against logic
      const existingById = new Map<string, ChecklistTemplate>();
      (rawRows ?? []).forEach((r) => existingById.set(r.id, r));

      // Track which existing IDs are still present in the working copy — any
      // that aren't get deleted from SP.
      const keptIds = new Set<string>();

      for (let i = 0; i < cleaned.length; i++) {
        const row = cleaned[i];
        const fields = itemToTemplateFields(row, i);
        if (row.rowId && existingById.has(row.rowId)) {
          keptIds.add(row.rowId);
          await updateListItem(LIST_NAMES.ChecklistTemplates, row.rowId, fields);
        } else {
          // New row — create
          await createListItem(LIST_NAMES.ChecklistTemplates, fields);
        }
      }

      // Delete any SP rows the user removed
      for (const [id] of existingById) {
        if (!keptIds.has(id)) {
          await deleteListItem(LIST_NAMES.ChecklistTemplates, id);
        }
      }

      // Refresh state from SP so rowIds for new rows are populated
      refetch?.();
      setSavedStamp(new Date().toISOString());
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleResetToDefaults = () => {
    if (!confirm("Replace the current list with the built-in defaults? Your custom changes will be removed (use Export JSON first if you want a backup).")) return;
    setItems(DOR_FILING_CHECKLIST.map((t) => ({ ...t, rowId: null })));
    // User still needs to click Save to push the defaults up to SharePoint
  };

  const handleImportFromBrowser = async () => {
    const override = readLocalChecklistOverride();
    if (!override || override.length === 0) {
      setImportError('No saved configuration found in this browser.');
      return;
    }
    setSaving(true);
    setSaveError(null);
    setImportError(null);
    try {
      // Wipe existing SP rows first so we don't double up, then create from the local override
      for (const row of rawRows ?? []) {
        await deleteListItem(LIST_NAMES.ChecklistTemplates, row.id);
      }
      for (let i = 0; i < override.length; i++) {
        await createListItem(LIST_NAMES.ChecklistTemplates, itemToTemplateFields(override[i], i));
      }
      clearLocalChecklistOverride();
      refetch?.();
      setSavedStamp(new Date().toISOString());
      setShowImportConfirm(false);
    } catch (err) {
      setImportError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleExport = () => {
    const blob = new Blob([JSON.stringify(items.map(({ rowId: _id, ...rest }) => rest), null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const stamp = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `cahp-checklist-templates-${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleImportClick = () => {
    setImportError(null);
    fileInputRef.current?.click();
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (!Array.isArray(parsed)) throw new Error('JSON root must be an array');
      const validated: EditorRow[] = parsed.map((row, idx) => {
        if (!row || typeof row !== 'object') throw new Error(`Row ${idx} is not an object`);
        if (typeof row.title !== 'string') throw new Error(`Row ${idx} missing title`);
        if (typeof row.category !== 'string') throw new Error(`Row ${idx} missing category`);
        if (row.scope !== 'cahp' && row.scope !== 'owner' && row.scope !== 'property') {
          throw new Error(`Row ${idx} has invalid scope: ${row.scope}`);
        }
        if (row.state !== undefined && row.state !== 'SC' && row.state !== 'NC') {
          throw new Error(`Row ${idx} has invalid state: ${row.state}`);
        }
        return {
          rowId: null,
          title: row.title,
          category: row.category as ItemCategory,
          scope: row.scope as FilingChecklistScope,
          notes: typeof row.notes === 'string' ? row.notes : undefined,
          library: typeof row.library === 'string' ? row.library : undefined,
          state: row.state as CahpState | undefined,
        };
      });
      setItems(validated);
      setImportError(null);
    } catch (err) {
      setImportError(err instanceof Error ? err.message : String(err));
    } finally {
      e.target.value = '';
    }
  };

  // Detect unsaved changes — compare working copy to the live SP rows
  const dirty = useMemo(() => {
    if (loading) return false;
    const existingMap = new Map<string, ChecklistTemplate>();
    (rawRows ?? []).forEach((r) => existingMap.set(r.id, r));
    // Different row count => dirty
    if (items.length !== (rawRows?.length ?? 0)) return true;
    for (let i = 0; i < items.length; i++) {
      const row = items[i];
      if (!row.rowId) return true;
      const sp = existingMap.get(row.rowId);
      if (!sp) return true;
      const spFields = sp.fields;
      if (
        spFields.Title !== row.title ||
        spFields.TemplateCategory !== row.category ||
        spFields.TemplateScope !== row.scope ||
        (spFields.TemplateNotes ?? undefined) !== (row.notes || undefined) ||
        (spFields.TemplateLibrary ?? undefined) !== (row.library || undefined) ||
        spFields.TemplateState !== row.state ||
        (spFields.TemplateSortOrder ?? -1) !== i
      ) {
        return true;
      }
    }
    return false;
  }, [items, rawRows, loading, savedStamp]);

  return (
    <div className="space-y-4">
      {error ? (
        <div className="bg-red-50 border border-red-200 rounded-md p-3 text-xs text-red-900">
          <strong>Couldn't reach the Checklist Templates list.</strong>
          <div className="mt-1">
            {error.message}
          </div>
          <div className="mt-2">
            If you haven't provisioned the list yet, run{' '}
            <code className="bg-white px-1 rounded">scripts\provision-checklist-templates.ps1</code> from
            the repo root. The app is falling back to the hardcoded DOR defaults below.
          </div>
        </div>
      ) : usingFallback ? (
        <div className="bg-blue-50 border border-blue-200 rounded-md p-3 text-xs text-blue-900">
          <strong>Showing built-in defaults.</strong> The Checklist Templates list in SharePoint is empty
          (or hasn't been provisioned). Edit the list below and click <strong>Save changes</strong> to
          push these defaults up so the rest of the team sees them.
        </div>
      ) : (
        <div className="bg-emerald-50 border border-emerald-200 rounded-md p-3 text-xs text-emerald-900">
          <strong>Synced to SharePoint.</strong> Edits made here are visible to every teammate. No more
          per-browser configuration.
        </div>
      )}

      {localOverride && (
        <div className="bg-amber-50 border border-amber-200 rounded-md p-3 text-xs text-amber-900 flex items-start justify-between gap-3 flex-wrap">
          <div>
            <strong>Older browser-only config detected.</strong> You have {localOverride.length} item{localOverride.length === 1 ? '' : 's'} saved in this browser's localStorage from the previous version.
            Click <strong>Import from this browser</strong> to push them up to SharePoint (this replaces the current SharePoint list with your saved items, then clears the local copy).
          </div>
          <button
            onClick={() => setShowImportConfirm(true)}
            disabled={saving}
            className="text-xs px-2.5 py-1.5 rounded-md bg-amber-700 hover:bg-amber-800 text-white font-medium disabled:opacity-50 flex-shrink-0"
          >
            Import from this browser
          </button>
        </div>
      )}

      {showImportConfirm && (
        <div className="bg-white border-2 border-amber-400 rounded-md p-3 text-xs">
          <p className="text-amber-900 font-medium mb-2">
            Replace SharePoint's {rawRows?.length ?? 0} item{(rawRows?.length ?? 0) === 1 ? '' : 's'} with the {localOverride?.length ?? 0} item{(localOverride?.length ?? 0) === 1 ? '' : 's'} saved in this browser?
          </p>
          <div className="flex items-center gap-2">
            <button
              onClick={handleImportFromBrowser}
              disabled={saving}
              className="px-3 py-1.5 rounded-md bg-amber-700 hover:bg-amber-800 text-white font-medium disabled:opacity-50"
            >
              {saving ? 'Importing…' : 'Yes, import + clear local copy'}
            </button>
            <button
              onClick={() => setShowImportConfirm(false)}
              disabled={saving}
              className="px-3 py-1.5 rounded-md border border-gray-300 hover:bg-gray-50 disabled:opacity-50"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      <div className="bg-white border border-gray-200 rounded-lg shadow-card">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between gap-2 flex-wrap">
          <div>
            <h3 className="text-sm font-semibold text-teal-900">Filing Checklist Templates</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {items.length} item{items.length === 1 ? '' : 's'}
              {dirty && <span className="ml-2 text-amber-700 font-semibold">· unsaved changes</span>}
            </p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={handleImportClick}
              className="text-xs px-2.5 py-1.5 rounded-md border border-gray-300 bg-white hover:bg-gray-50 inline-flex items-center gap-1.5"
            >
              <Icon name="folder" size={12} />
              Import JSON
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept="application/json,.json"
              onChange={handleImportFile}
              className="hidden"
            />
            <button
              onClick={handleExport}
              className="text-xs px-2.5 py-1.5 rounded-md border border-gray-300 bg-white hover:bg-gray-50 inline-flex items-center gap-1.5"
            >
              <Icon name="file" size={12} />
              Export JSON
            </button>
            <button
              onClick={handleResetToDefaults}
              className="text-xs px-2.5 py-1.5 rounded-md border border-gray-300 bg-white hover:bg-gray-50 text-gray-700"
              title="Replace the current working copy with the hardcoded defaults (click Save to push to SharePoint)"
            >
              Load defaults
            </button>
            <button
              onClick={handleSave}
              disabled={!dirty || saving}
              className="text-xs px-3 py-1.5 rounded-md font-medium bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white inline-flex items-center gap-1.5"
            >
              {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
              {saving ? 'Saving…' : 'Save changes'}
            </button>
          </div>
        </div>

        {(saveError || importError) && (
          <div className="px-4 py-2 bg-red-50 border-b border-red-200 text-xs text-red-800">
            {saveError ?? importError}
          </div>
        )}

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-[11px] font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-3 py-2 text-left w-8"></th>
                <th className="px-3 py-2 text-left">Title</th>
                <th className="px-3 py-2 text-left w-44">Category</th>
                <th className="px-3 py-2 text-left w-36">Scope</th>
                <th className="px-3 py-2 text-left w-32">State</th>
                <th className="px-3 py-2 text-left w-48">SharePoint Library</th>
                <th className="px-3 py-2 text-right w-24">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {items.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-3 py-6 text-center text-xs text-gray-500 italic">
                    No items. Add one below or click <strong>Load defaults</strong> to seed the hardcoded list.
                  </td>
                </tr>
              ) : (
                items.map((item, idx) => (
                  <tr key={item.rowId ?? `new-${idx}`} className="hover:bg-gray-50">
                    <td className="px-3 py-2 align-top">
                      <div className="flex flex-col gap-0.5 text-gray-400">
                        <button
                          onClick={() => moveItem(idx, -1)}
                          disabled={idx === 0 || saving}
                          className="hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed text-[10px] leading-none"
                          title="Move up"
                        >
                          ▲
                        </button>
                        <button
                          onClick={() => moveItem(idx, 1)}
                          disabled={idx === items.length - 1 || saving}
                          className="hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed text-[10px] leading-none"
                          title="Move down"
                        >
                          ▼
                        </button>
                      </div>
                    </td>
                    <td className="px-3 py-2 align-top">
                      <input
                        type="text"
                        value={item.title}
                        onChange={(e) => updateItem(idx, { title: e.target.value })}
                        disabled={saving}
                        placeholder="Item title (e.g., 'CAHP Operating Agreement')"
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500"
                      />
                      <textarea
                        value={item.notes ?? ''}
                        onChange={(e) => updateItem(idx, { notes: e.target.value })}
                        disabled={saving}
                        placeholder="Notes (optional)"
                        rows={2}
                        className="w-full mt-1 px-2 py-1 border border-gray-200 rounded text-[11px] text-gray-700 resize-none focus:outline-none focus:border-teal-500"
                      />
                    </td>
                    <td className="px-3 py-2 align-top">
                      <select
                        value={item.category}
                        onChange={(e) => updateItem(idx, { category: e.target.value as ItemCategory })}
                        disabled={saving}
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs bg-white focus:outline-none focus:border-teal-500"
                      >
                        {ITEM_CATEGORIES.map((c) => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-3 py-2 align-top">
                      <select
                        value={item.scope}
                        onChange={(e) => updateItem(idx, { scope: e.target.value as FilingChecklistScope })}
                        disabled={saving}
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs bg-white focus:outline-none focus:border-teal-500"
                      >
                        {SCOPES.map((s) => (
                          <option key={s.value} value={s.value} title={s.help}>{s.label}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-3 py-2 align-top">
                      <select
                        value={item.state ?? ''}
                        onChange={(e) => updateItem(idx, { state: (e.target.value || undefined) as CahpState | undefined })}
                        disabled={saving}
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs bg-white focus:outline-none focus:border-teal-500"
                      >
                        {STATE_OPTIONS.map((s) => (
                          <option key={s.value} value={s.value}>{s.label}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-3 py-2 align-top">
                      <select
                        value={item.library ?? ''}
                        onChange={(e) => updateItem(idx, { library: e.target.value || undefined })}
                        disabled={saving}
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs bg-white focus:outline-none focus:border-teal-500"
                      >
                        {LIBRARY_OPTIONS.map((l) => (
                          <option key={l.value} value={l.value}>{l.label}</option>
                        ))}
                      </select>
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
                ))
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
            Add checklist item
          </button>
        </div>
      </div>

      <div className="text-[11px] text-gray-500 leading-relaxed">
        <strong>How it's used:</strong> When you generate a filing checklist for a property or submittal, the system
        walks this list, auto-matches existing documents in the SharePoint libraries, and creates Outstanding Items
        for anything missing. The library column controls where the auto-match looks — leave it on{' '}
        <em>Use category default</em> unless an item belongs in a non-standard library.
      </div>
    </div>
  );
}
