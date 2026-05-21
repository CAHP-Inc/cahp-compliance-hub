import { useMemo, useState, useRef } from 'react';
import {
  DOR_FILING_CHECKLIST,
  getFilingChecklist,
  saveFilingChecklist,
  resetFilingChecklist,
  hasCustomFilingChecklist,
  type FilingChecklistItem,
  type FilingChecklistScope,
} from '../lib/filing-checklist';
import { PROPERTY_LINKED_LIBRARIES } from './UploadDocumentModal';
import { Icon } from './ui/Icon';
import type { ItemCategory, CahpState } from '../lib/sharepoint';

/**
 * Settings → Checklist Templates editor.
 *
 * Edits the list of items the Filing Checklist Generator creates as
 * outstanding items when run against a property/submittal. Persistence is
 * browser-local (localStorage). Use Export JSON / Import JSON to share a
 * configuration across browsers or teammates.
 */

// Mirrors ItemCategory (kept in sync manually to expose as choice list)
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

export function ChecklistTemplatesEditor() {
  // Snapshot active list once on mount; mutating local state, saving on Save click.
  const [items, setItems] = useState<FilingChecklistItem[]>(() => getFilingChecklist());
  const [savedStamp, setSavedStamp] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [importError, setImportError] = useState<string | null>(null);

  const isCustomized = useMemo(() => hasCustomFilingChecklist(), [savedStamp]);

  const dirty = useMemo(() => {
    const baseline = getFilingChecklist();
    return JSON.stringify(baseline) !== JSON.stringify(items);
  }, [items, savedStamp]);

  const updateItem = (idx: number, patch: Partial<FilingChecklistItem>) => {
    setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, ...patch } : it)));
  };

  const removeItem = (idx: number) => {
    setItems((prev) => prev.filter((_, i) => i !== idx));
  };

  const addItem = () => {
    setItems((prev) => [
      ...prev,
      { title: '', category: 'Other', scope: 'property', notes: '' },
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

  const handleSave = () => {
    // Drop rows the user left blank
    const cleaned = items.filter((i) => i.title.trim().length > 0);
    saveFilingChecklist(cleaned);
    setItems(cleaned);
    setSavedStamp(new Date().toISOString());
  };

  const handleReset = () => {
    if (!confirm('Reset to the built-in defaults? Your custom changes will be lost (use Export first if you want a backup).')) return;
    resetFilingChecklist();
    setItems(DOR_FILING_CHECKLIST);
    setSavedStamp(new Date().toISOString());
  };

  const handleExport = () => {
    const blob = new Blob([JSON.stringify(items, null, 2)], { type: 'application/json' });
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
      const validated: FilingChecklistItem[] = parsed.map((row, idx) => {
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
      // Reset so same file can be re-imported after fix
      e.target.value = '';
    }
  };

  return (
    <div className="space-y-4">
      <div className="bg-amber-50 border border-amber-200 rounded-md p-3 text-xs text-amber-900">
        <strong>Storage:</strong> Edits live in this browser (localStorage). To share a configuration across browsers
        or teammates, use <strong>Export JSON</strong> here and <strong>Import JSON</strong> elsewhere. We can promote
        this to a shared SharePoint list when multi-user sync becomes important.
      </div>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between gap-2 flex-wrap">
          <div>
            <h3 className="text-sm font-semibold text-teal-900">Filing Checklist Templates</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {items.length} item{items.length === 1 ? '' : 's'}
              {isCustomized ? ' · custom configuration' : ' · using built-in defaults'}
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
              onClick={handleReset}
              className="text-xs px-2.5 py-1.5 rounded-md border border-gray-300 bg-white hover:bg-gray-50 text-gray-700"
              title="Discard custom edits and restore the hardcoded default list"
            >
              Reset to defaults
            </button>
            <button
              onClick={handleSave}
              disabled={!dirty}
              className="text-xs px-3 py-1.5 rounded-md font-medium bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white"
            >
              Save changes
            </button>
          </div>
        </div>

        {importError && (
          <div className="px-4 py-2 bg-red-50 border-b border-red-200 text-xs text-red-800">
            Import failed: {importError}
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
                    No items. Add one below or click <strong>Reset to defaults</strong> to load the built-in list.
                  </td>
                </tr>
              ) : (
                items.map((item, idx) => (
                  <tr key={idx} className="hover:bg-gray-50">
                    <td className="px-3 py-2 align-top">
                      <div className="flex flex-col gap-0.5 text-gray-400">
                        <button
                          onClick={() => moveItem(idx, -1)}
                          disabled={idx === 0}
                          className="hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed text-[10px] leading-none"
                          title="Move up"
                        >
                          ▲
                        </button>
                        <button
                          onClick={() => moveItem(idx, 1)}
                          disabled={idx === items.length - 1}
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
                        placeholder="Item title (e.g., 'CAHP Operating Agreement')"
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500"
                      />
                      <textarea
                        value={item.notes ?? ''}
                        onChange={(e) => updateItem(idx, { notes: e.target.value })}
                        placeholder="Notes (optional) — explains what the doc is and where it lives"
                        rows={2}
                        className="w-full mt-1 px-2 py-1 border border-gray-200 rounded text-[11px] text-gray-700 resize-none focus:outline-none focus:border-teal-500"
                      />
                    </td>
                    <td className="px-3 py-2 align-top">
                      <select
                        value={item.category}
                        onChange={(e) => updateItem(idx, { category: e.target.value as ItemCategory })}
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
                        className="w-full px-2 py-1 border border-gray-300 rounded text-xs bg-white focus:outline-none focus:border-teal-500"
                        title="When set, this item is only added for properties in that state"
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
                        className="text-[11px] text-error hover:text-red-700 font-medium px-2 py-1 rounded hover:bg-red-50"
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
            className="text-xs px-3 py-1.5 rounded-md border border-dashed border-teal-400 text-teal-700 hover:bg-teal-50 inline-flex items-center gap-1.5"
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
