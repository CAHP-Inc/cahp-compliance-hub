import { useEffect, useMemo, useState } from 'react';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type EmailTemplate,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';

/**
 * Settings → Email Templates editor.
 *
 * Stored in the shared SharePoint `Email Templates` list, same pattern as
 * Checklist Templates: teammates see the same set.
 *
 * Variable substitution at send time:
 *   {{contact}}, {{contact_email}}, {{property}}, {{properties}},
 *   {{owner}}, {{date}}, {{user}}, {{user_email}}
 */

interface EditorRow {
  rowId: string | null;
  title: string;
  subject: string;
  body: string;
  notes: string;
}

const DEFAULT_TEMPLATES: Omit<EditorRow, 'rowId'>[] = [
  {
    title: 'Document Request',
    subject: 'CAHP: Documents needed for {{property}}',
    body:
      "Hi {{contact}},\n\n" +
      "I'm following up on the documents we need to complete the CAHP filing for {{property}}.\n\n" +
      "Could you send the following at your earliest convenience:\n" +
      "  • [list items here]\n\n" +
      "Let me know if you have any questions.\n\n" +
      "Thanks,\n{{user}}",
    notes: 'Use when you need missing documents from a property owner contact.',
  },
  {
    title: 'Status Update',
    subject: 'CAHP filing status — {{property}}',
    body:
      "Hi {{contact}},\n\n" +
      "Quick update on the CAHP filing for {{property}}:\n\n" +
      "  • [current status]\n" +
      "  • [next steps]\n\n" +
      "Let me know if you have any questions.\n\n" +
      "Best,\n{{user}}",
    notes: 'Periodic update to owner contacts on filing progress.',
  },
  {
    title: 'DOR Inquiry',
    subject: 'Question about CAHP filing process',
    body:
      "Hello,\n\n" +
      "I have a question about [topic] for our CAHP filings.\n\n" +
      "[question details]\n\n" +
      "Thanks for your help,\n{{user}}\n{{user_email}}",
    notes: 'Use when contacting DOR with a procedural question.',
  },
];

export function EmailTemplatesEditor() {
  const list = useSharePointList<EmailTemplate>(LIST_NAMES.EmailTemplates, { top: 500 });
  const [items, setItems] = useState<EditorRow[]>([]);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  // Hydrate working copy from SharePoint
  useEffect(() => {
    if (list.loading) return;
    if (list.data && list.data.length > 0) {
      const sorted = [...list.data].sort((a, b) => {
        const aOrder = a.fields.TemplateSortOrder ?? Number.MAX_SAFE_INTEGER;
        const bOrder = b.fields.TemplateSortOrder ?? Number.MAX_SAFE_INTEGER;
        return aOrder - bOrder;
      });
      setItems(
        sorted.map((row) => ({
          rowId: row.id,
          title: row.fields.Title ?? '',
          subject: row.fields.TemplateSubject ?? '',
          body: row.fields.TemplateBody ?? '',
          notes: row.fields.TemplateNotes ?? '',
        })),
      );
    } else if (!list.data) {
      // List doesn't exist yet — leave the editor empty so the user can
      // see the provisioning hint banner above the table.
      setItems([]);
    } else {
      // List exists but empty — seed with sensible starter templates so
      // the user has something to work with on the first visit.
      setItems(DEFAULT_TEMPLATES.map((t) => ({ rowId: null, ...t })));
    }
  }, [list.loading, list.data]);

  const usingFallback = !list.loading && list.data?.length === 0;

  const updateItem = (idx: number, patch: Partial<EditorRow>) => {
    setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, ...patch } : it)));
  };
  const removeItem = (idx: number) => setItems((prev) => prev.filter((_, i) => i !== idx));
  const addItem = () =>
    setItems((prev) => [
      ...prev,
      { rowId: null, title: '', subject: '', body: '', notes: '' },
    ]);
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
      const existingById = new Map<string, EmailTemplate>();
      (list.data ?? []).forEach((r) => existingById.set(r.id, r));
      const keptIds = new Set<string>();

      for (let i = 0; i < cleaned.length; i++) {
        const row = cleaned[i];
        const fields = {
          Title: row.title,
          TemplateSubject: row.subject || null,
          TemplateBody: row.body || null,
          TemplateNotes: row.notes || null,
          TemplateSortOrder: i,
        };
        if (row.rowId && existingById.has(row.rowId)) {
          keptIds.add(row.rowId);
          await updateListItem(LIST_NAMES.EmailTemplates, row.rowId, fields);
        } else {
          await createListItem(LIST_NAMES.EmailTemplates, fields);
        }
      }
      for (const [id] of existingById) {
        if (!keptIds.has(id)) await deleteListItem(LIST_NAMES.EmailTemplates, id);
      }
      list.refetch?.();
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const dirty = useMemo(() => {
    if (list.loading) return false;
    const existingMap = new Map<string, EmailTemplate>();
    (list.data ?? []).forEach((r) => existingMap.set(r.id, r));
    if (items.length !== (list.data?.length ?? 0)) return true;
    for (let i = 0; i < items.length; i++) {
      const row = items[i];
      if (!row.rowId) return true;
      const sp = existingMap.get(row.rowId);
      if (!sp) return true;
      const f = sp.fields;
      if (
        f.Title !== row.title ||
        (f.TemplateSubject ?? '') !== row.subject ||
        (f.TemplateBody ?? '') !== row.body ||
        (f.TemplateNotes ?? '') !== row.notes ||
        (f.TemplateSortOrder ?? -1) !== i
      ) {
        return true;
      }
    }
    return false;
  }, [items, list.data, list.loading]);

  return (
    <div className="space-y-4">
      {list.error ? (
        <div className="bg-red-50 border border-red-200 rounded-md p-3 text-xs text-red-900">
          <strong>Couldn't reach the Email Templates list.</strong>
          <div className="mt-1">{list.error.message}</div>
          <div className="mt-2">
            Run <code className="bg-white px-1 rounded">scripts\provision-email-templates.ps1</code>{' '}
            to create the list. The Compose Email modal will still work without templates.
          </div>
        </div>
      ) : usingFallback ? (
        <div className="bg-blue-50 border border-blue-200 rounded-md p-3 text-xs text-blue-900">
          <strong>Showing starter templates.</strong> The Email Templates list is empty; edit and click{' '}
          <strong>Save changes</strong> to share these with the team.
        </div>
      ) : (
        <div className="bg-emerald-50 border border-emerald-200 rounded-md p-3 text-xs text-emerald-900">
          <strong>Synced to SharePoint.</strong> Templates appear in the Compose Email modal's template picker.
        </div>
      )}

      <div className="bg-gray-50 border border-gray-200 rounded-md p-3 text-[11px] text-gray-700">
        <strong>Variables you can use in subject + body:</strong>{' '}
        <code className="bg-white px-1 rounded">{'{{contact}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{contact_email}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{property}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{properties}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{owner}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{date}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{user}}'}</code>,{' '}
        <code className="bg-white px-1 rounded">{'{{user_email}}'}</code>.{' '}
        Unknown variables pass through unchanged.
      </div>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between gap-2 flex-wrap">
          <div>
            <h3 className="text-sm font-semibold text-teal-900">Email Templates</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {items.length} template{items.length === 1 ? '' : 's'}
              {dirty && <span className="ml-2 text-amber-700 font-semibold">· unsaved changes</span>}
            </p>
          </div>
          <button
            onClick={handleSave}
            disabled={!dirty || saving}
            className="text-xs px-3 py-1.5 rounded-md font-medium bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white inline-flex items-center gap-1.5"
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

        <div className="divide-y divide-gray-100">
          {items.length === 0 ? (
            <div className="px-4 py-6 text-center text-xs text-gray-500 italic">
              No templates yet. Click <strong>Add template</strong> below.
            </div>
          ) : (
            items.map((item, idx) => (
              <div key={item.rowId ?? `new-${idx}`} className="px-4 py-3 space-y-2">
                <div className="flex items-start gap-2">
                  <div className="flex flex-col gap-0.5 text-gray-400 pt-2">
                    <button
                      onClick={() => moveItem(idx, -1)}
                      disabled={idx === 0 || saving}
                      className="hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed text-[10px] leading-none"
                      title="Move up"
                    >▲</button>
                    <button
                      onClick={() => moveItem(idx, 1)}
                      disabled={idx === items.length - 1 || saving}
                      className="hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed text-[10px] leading-none"
                      title="Move down"
                    >▼</button>
                  </div>
                  <input
                    type="text"
                    value={item.title}
                    onChange={(e) => updateItem(idx, { title: e.target.value })}
                    disabled={saving}
                    placeholder="Template name (e.g., 'Document Request')"
                    className="flex-1 px-2 py-1 border border-gray-300 rounded text-sm font-medium focus:outline-none focus:border-teal-500"
                  />
                  <button
                    onClick={() => removeItem(idx)}
                    disabled={saving}
                    className="text-[11px] text-error hover:text-red-700 font-medium px-2 py-1 rounded hover:bg-red-50 disabled:opacity-50"
                  >
                    Remove
                  </button>
                </div>
                <input
                  type="text"
                  value={item.subject}
                  onChange={(e) => updateItem(idx, { subject: e.target.value })}
                  disabled={saving}
                  placeholder="Subject line (variables OK)"
                  className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500"
                />
                <textarea
                  value={item.body}
                  onChange={(e) => updateItem(idx, { body: e.target.value })}
                  disabled={saving}
                  rows={6}
                  placeholder="Email body — variables OK"
                  className="w-full px-2 py-1 border border-gray-300 rounded text-xs font-mono focus:outline-none focus:border-teal-500"
                />
                <input
                  type="text"
                  value={item.notes}
                  onChange={(e) => updateItem(idx, { notes: e.target.value })}
                  disabled={saving}
                  placeholder="Internal notes (when to use this template — not sent)"
                  className="w-full px-2 py-1 border border-gray-200 rounded text-[11px] text-gray-600 focus:outline-none focus:border-teal-500"
                />
              </div>
            ))
          )}
        </div>

        <div className="px-4 py-3 border-t border-gray-100 bg-gray-50 rounded-b-lg">
          <button
            onClick={addItem}
            disabled={saving}
            className="text-xs px-3 py-1.5 rounded-md border border-dashed border-teal-400 text-teal-700 hover:bg-teal-50 inline-flex items-center gap-1.5 disabled:opacity-50"
          >
            <Icon name="plus" size={12} />
            Add template
          </button>
        </div>
      </div>
    </div>
  );
}
