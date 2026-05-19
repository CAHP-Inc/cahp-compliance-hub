import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type OutstandingItem,
  type OutstandingItemFields,
  type Property,
  type ItemStatus,
  type ItemPriority,
  type ItemCategory,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { LinkOrUploadDocumentModal } from '../components/LinkOrUploadDocumentModal';
import { AssigneePicker } from '../components/AssigneePicker';
import { formatDateOnly } from '../lib/dates';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';

const STATUSES: ItemStatus[] = ['Not Started', 'In Progress', 'Blocked', 'Done', 'Not Applicable'];
const PRIORITIES: ItemPriority[] = ['Critical', 'High', 'Medium', 'Low'];
const CATEGORIES: ItemCategory[] = [
  'Operating Agreement',
  'LURA',
  'AMI Certification',
  'Org Chart',
  'Income Documentation',
  'Signed Submittal',
  'Determination Letter',
  'Other',
];

const STATUS_STYLES: Record<ItemStatus, string> = {
  'Not Started': 'bg-gray-100 text-gray-800',
  'In Progress': 'bg-blue-100 text-blue-800',
  'Blocked': 'bg-red-100 text-red-800',
  'Done': 'bg-green-100 text-green-800',
  'Requested': 'bg-gray-100 text-gray-800',
  'Overdue': 'bg-amber-100 text-amber-800',
  'Received': 'bg-green-100 text-green-800',
  'Not Applicable': 'bg-gray-100 text-gray-500',
};

const PRIORITY_STYLES: Record<ItemPriority, string> = {
  Critical: 'bg-red-100 text-red-800',
  High: 'bg-amber-100 text-amber-800',
  Medium: 'bg-blue-100 text-blue-800',
  Low: 'bg-gray-100 text-gray-600',
};

export function OutstandingItemDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const { data: item, loading, error, refetch } = useSharePointItem<OutstandingItem>(
    LIST_NAMES.Outstanding,
    id
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<OutstandingItemFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [linkOrUploadOpen, setLinkOrUploadOpen] = useState(false);

  useEffect(() => {
    if (item && !editing) setDraft({ ...item.fields });
  }, [item?.id, item?.lastModifiedDateTime, editing]);

  const property = useMemo(() => {
    if (!item || !properties.data || !item.fields.PropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(item.fields.PropertyLookupId)) ?? null;
  }, [item, properties.data]);

  const overdue = useMemo(() => {
    if (!item || !item.fields.DueDate) return false;
    const due = new Date(item.fields.DueDate);
    const closed = item.fields.ItemStatus === 'Done' || item.fields.ItemStatus === 'Received' || item.fields.ItemStatus === 'Not Applicable';
    return due.getTime() < Date.now() && !closed;
  }, [item]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading item…</span>
        </div>
      </div>
    );
  }

  if (error || !item || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Outstanding Items" parentTo="/outstanding-items" currentLabel="Item Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load item</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : item.fields;

  const handleFieldChange = <K extends keyof OutstandingItemFields>(
    field: K,
    value: OutstandingItemFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...item.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...item.fields });
    setSaveError(null);
    setEditing(false);
  };

  const handleSave = async () => {
    if (!draft) return;
    setSaving(true);
    setSaveError(null);
    try {
      const changed: Record<string, unknown> = {};
      Object.keys(draft).forEach((key) => {
        const k = key as keyof OutstandingItemFields;
        if (draft[k] !== item.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      // If status changed to Done, set DateReceivedItem
      if (changed.ItemStatus === 'Done' || changed.ItemStatus === 'Received') {
        if (!item.fields.DateReceivedItem) {
          changed.DateReceivedItem = new Date().toISOString();
        }
      }
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Outstanding, item.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Permanently delete "${item.fields.Title}"?\n\nThis cannot be undone.`)) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Outstanding, item.id);
      navigate('/outstanding-items');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const handleQuickStatusChange = async (newStatus: ItemStatus) => {
    setSaving(true);
    setSaveError(null);
    try {
      const patch: Record<string, unknown> = { ItemStatus: newStatus };
      if (newStatus === 'Done' && !item.fields.DateReceivedItem) {
        patch.DateReceivedItem = new Date().toISOString();
      }
      await updateListItem(LIST_NAMES.Outstanding, item.id, patch);
      await refetch();
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Outstanding Items" parentTo="/outstanding-items" currentLabel={item.fields.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1 flex-wrap">
            <h1 className="text-2xl font-bold text-teal-700">{item.fields.Title}</h1>
            {item.fields.ItemStatus && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[item.fields.ItemStatus]}`}>
                {item.fields.ItemStatus}
              </span>
            )}
            {item.fields.Priority && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${PRIORITY_STYLES[item.fields.Priority]}`}>
                {item.fields.Priority}
              </span>
            )}
            {overdue && (
              <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-red-100 text-red-800 border border-red-200">
                ⚠ Overdue
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500">
            {property ? (
              <Link to={`/properties/${property.id}`} className="text-teal-700 hover:text-teal-900 underline">
                {property.fields.Title}
              </Link>
            ) : (
              <span className="italic text-gray-400">no property linked</span>
            )}
            {item.fields.DueDate && (
              <>
                {' · '}Due{' '}
                <span className={overdue ? 'text-error font-semibold' : ''}>
                  {formatDateOnly(item.fields.DueDate)}
                </span>
              </>
            )}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              <button
                onClick={handleDelete}
                disabled={saving}
                className="px-3 py-1.5 border border-red-300 text-error hover:bg-red-50 rounded-md text-sm font-medium disabled:opacity-50"
              >
                Delete
              </button>
              <button
                onClick={() => setLinkOrUploadOpen(true)}
                className="px-3 py-1.5 bg-gold-500 hover:bg-gold-600 text-teal-900 rounded-md text-sm font-medium flex items-center gap-1.5"
              >
                <Icon name="folder" size={14} />
                {item.fields.RelatedDocUrl ? 'Replace Document' : 'Link / Upload Document'}
              </button>
              <button
                onClick={handleEdit}
                className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5"
              >
                <Icon name="settings" size={14} />
                Edit
              </button>
            </>
          )}
          {editing && <EditingActionButtons saving={saving} onCancel={handleCancel} onSave={handleSave} />}
        </div>
      </div>

      <SaveErrorBanner error={saveError} />

      {/* Quick action: status pipeline */}
      {!editing && item.fields.ItemStatus !== 'Done' && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-4">
          <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Quick Status</div>
          <div className="flex flex-wrap gap-2">
            {(['Not Started', 'In Progress', 'Blocked', 'Done'] as ItemStatus[]).map((s) => {
              const isCurrent = item.fields.ItemStatus === s || (s === 'Not Started' && item.fields.ItemStatus === 'Requested');
              return (
                <button
                  key={s}
                  onClick={() => handleQuickStatusChange(s)}
                  disabled={saving || isCurrent}
                  className={`px-3 py-1 rounded-md text-xs font-medium transition-colors disabled:opacity-50 ${
                    isCurrent
                      ? `${STATUS_STYLES[s]} cursor-default`
                      : s === 'Done'
                        ? 'bg-green-600 hover:bg-green-700 text-white'
                        : 'bg-white border border-gray-300 text-gray-700 hover:bg-gray-50'
                  }`}
                >
                  {isCurrent ? `✓ ${s}` : `→ ${s}`}
                </button>
              );
            })}
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Section title="Item Details">
          <EditableField
            label="Title"
            value={display.Title}
            editing={editing}
            onChange={(v) => handleFieldChange('Title', v as string)}
            required
          />
          <EditableField
            label="Status"
            value={display.ItemStatus}
            editing={editing}
            type="choice"
            choices={STATUSES}
            onChange={(v) => handleFieldChange('ItemStatus', v as ItemStatus)}
          />
          <EditableField
            label="Priority"
            value={display.Priority}
            editing={editing}
            type="choice"
            choices={PRIORITIES}
            onChange={(v) => handleFieldChange('Priority', v as ItemPriority)}
          />
          <EditableField
            label="Category"
            value={display.ItemCategory}
            editing={editing}
            type="choice"
            choices={CATEGORIES}
            onChange={(v) => handleFieldChange('ItemCategory', v as ItemCategory)}
          />
          {editing ? (
            <div className="flex items-start gap-3">
              <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">Assigned To</label>
              <div className="flex-1">
                <AssigneePicker
                  value={display.AssignedTo}
                  onChange={(v) => handleFieldChange('AssignedTo', v)}
                />
                <p className="text-[11px] text-gray-500 mt-1">
                  Pick a team member from the dropdown, or type a free-form name (vendor, owner, DOR, etc.)
                </p>
              </div>
            </div>
          ) : (
            <EditableField
              label="Assigned To"
              value={display.AssignedTo}
              editing={false}
              onChange={(v) => handleFieldChange('AssignedTo', v as string)}
            />
          )}
        </Section>

        <Section title="Dates & Tracking">
          <EditableField
            label="Due Date"
            value={display.DueDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DueDate', v as string)}
            mono
          />
          <EditableField
            label="Date Requested"
            value={display.DateRequested}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DateRequested', v as string)}
            mono
          />
          <EditableField
            label="Date Received"
            value={display.DateReceivedItem}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DateReceivedItem', v as string)}
            mono
          />
          <EditableField
            label="Follow-up Count"
            value={display.FollowUpCount?.toString()}
            editing={editing}
            type="number"
            onChange={(v) => handleFieldChange('FollowUpCount', v === '' ? undefined : Number(v))}
            mono
          />
          <EditableField
            label="Item Owner"
            value={display.ItemOwner}
            editing={editing}
            onChange={(v) => handleFieldChange('ItemOwner', v as string)}
          />
        </Section>

        <Section title="Notes" fullWidth>
          <EditableField
            label="Notes"
            value={display.ItemNotes}
            editing={editing}
            type="textarea"
            rows={5}
            hideLabel
            onChange={(v) => handleFieldChange('ItemNotes', v as string)}
          />
        </Section>
      </div>

      {/* Related Document panel — shows linked fulfilling doc if set */}
      {!editing && (item.fields.RelatedDocUrl || item.fields.RelatedDocFilename) && (
        <div className="mt-6 bg-green-50 border border-green-200 rounded-lg p-4">
          <div className="flex items-start gap-3">
            <Icon name="check" size={20} className="text-success flex-shrink-0 mt-0.5" />
            <div className="flex-1 min-w-0">
              <div className="text-xs font-semibold text-green-900 uppercase tracking-wider mb-1">
                Linked Document
              </div>
              {item.fields.RelatedDocUrl ? (
                <a
                  href={item.fields.RelatedDocUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-sm font-medium text-teal-700 hover:text-teal-900 hover:underline break-all inline-flex items-center gap-1.5"
                >
                  <Icon name="file" size={12} />
                  {item.fields.RelatedDocFilename ?? 'Open document'}
                </a>
              ) : (
                <div className="text-sm font-medium text-gray-900">
                  {item.fields.RelatedDocFilename}
                </div>
              )}
              {item.fields.RelatedDocLibrary && (
                <div className="text-[11px] text-gray-600 mt-1">Library: {item.fields.RelatedDocLibrary}</div>
              )}
            </div>
            <button
              onClick={() => setLinkOrUploadOpen(true)}
              className="px-2 py-1 text-xs text-teal-700 hover:text-teal-900 underline flex-shrink-0"
            >
              Replace
            </button>
          </div>
        </div>
      )}

      {linkOrUploadOpen && (
        <LinkOrUploadDocumentModal
          item={item}
          onClose={() => setLinkOrUploadOpen(false)}
          onSuccess={() => {
            setLinkOrUploadOpen(false);
            refetch();
          }}
        />
      )}
    </div>
  );
}
