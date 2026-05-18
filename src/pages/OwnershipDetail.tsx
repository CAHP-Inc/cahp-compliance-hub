import { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Ownership,
  type OwnershipFields,
  type RelationshipType,
  type Property,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
  formatDate,
} from '../components/detail';

const RELATIONSHIP_STYLES: Record<RelationshipType, string> = {
  'Managing Member': 'bg-teal-100 text-teal-800 border-teal-200',
  Member: 'bg-blue-100 text-blue-800 border-blue-200',
  Owner: 'bg-purple-100 text-purple-800 border-purple-200',
  Subsidiary: 'bg-amber-100 text-amber-800 border-amber-200',
  'Beneficial Owner': 'bg-pink-100 text-pink-800 border-pink-200',
};

// PR-09d gap #5 — per spec §5.5, every ownership UPDATE captures a reason
type OwnershipChangeReason = 'Buy-In' | 'Buy-Out' | 'Estate' | 'Initial Filing' | 'Other';

const CHANGE_REASONS: OwnershipChangeReason[] = [
  'Buy-In',
  'Buy-Out',
  'Estate',
  'Initial Filing',
  'Other',
];

const CHOICES = {
  RelationshipType: [
    'Managing Member', 'Member', 'Owner', 'Subsidiary', 'Beneficial Owner',
  ] as const,
} as const;

export function OwnershipDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<OwnershipFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  // PR-09d gap #5 — ownership changes require a reason per spec §5.5
  const [reasonModalOpen, setReasonModalOpen] = useState(false);
  const [reason, setReason] = useState<OwnershipChangeReason | ''>('');
  const [reasonNotes, setReasonNotes] = useState('');
  const [pendingChanges, setPendingChanges] = useState<Record<string, unknown> | null>(null);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);

  const { data: ownership, loading, error, refetch } = useSharePointItem<Ownership>(
    LIST_NAMES.Ownership, id
  );

  const { data: allProperties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 200 });
  const linkedProperty = ownership?.fields.LinkedPropertyLookupId
    ? allProperties?.find((p) => String(p.id) === String(ownership.fields.LinkedPropertyLookupId))
    : null;

  useEffect(() => {
    if (ownership && !editing) {
      setDraft({ ...ownership.fields });
    }
  }, [ownership?.id, ownership?.lastModifiedDateTime, editing]);

  const handleEdit = () => {
    if (!ownership) return;
    setDraft({ ...ownership.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    if (ownership) setDraft({ ...ownership.fields });
    setSaveError(null);
    setEditing(false);
  };

  const handleFieldChange = <K extends keyof OwnershipFields>(
    field: K, value: OwnershipFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleSave = async () => {
    if (!draft || !ownership) return;
    // Compute pending changes; if there are any, open the reason modal before writing
    const changes: Record<string, unknown> = {};
    (Object.keys(draft) as (keyof OwnershipFields)[]).forEach((k) => {
      const oldVal = ownership.fields[k];
      const newVal = draft[k];
      if (oldVal !== newVal) {
        changes[k as string] = newVal === '' ? null : newVal;
      }
    });
    if (Object.keys(changes).length === 0) {
      setEditing(false);
      return;
    }
    setPendingChanges(changes);
    setReason('');
    setReasonNotes('');
    setReasonModalOpen(true);
  };

  const confirmSave = async () => {
    if (!ownership || !pendingChanges) return;
    if (!reason) {
      setSaveError('Reason is required for ownership changes.');
      return;
    }
    setSaving(true);
    setSaveError(null);
    try {
      const fullReason = reasonNotes ? `${reason} — ${reasonNotes}` : reason;
      await updateListItem(LIST_NAMES.Ownership, ownership.id, pendingChanges, { reason: fullReason });
      await refetch();
      setEditing(false);
      setReasonModalOpen(false);
      setPendingChanges(null);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!ownership) return;
    setSaving(true);
    setSaveError(null);
    try {
      await deleteListItem(LIST_NAMES.Ownership, ownership.id);
      navigate('/ownership');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
      setShowDeleteConfirm(false);
    }
  };

  if (loading) {
    return (
      <div className="text-center py-12 text-gray-500">
        <div className="inline-flex items-center gap-3">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading ownership record…</span>
        </div>
      </div>
    );
  }

  if (error || !ownership || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Ownership" parentTo="/ownership" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Ownership record not found
          </div>
          <p className="text-sm text-red-700 mb-3">
            {error ? error.message : `No record with ID ${id} exists.`}
          </p>
          <button onClick={() => navigate('/ownership')} className="text-sm text-teal-700 hover:text-teal-900 font-medium underline">
            ← Back to Ownership Structure
          </button>
        </div>
      </div>
    );
  }

  const display = editing ? draft : ownership.fields;

  return (
    <div>
      <BreadcrumbBar
        parentLabel="Ownership"
        parentTo="/ownership"
        currentLabel={ownership.fields.Title}
      />

      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-3 flex-wrap">
            <h1 className="text-3xl font-bold text-teal-700">{ownership.fields.Title}</h1>
            {ownership.fields.RelationshipType && (
              <span className={`inline-block px-2.5 py-0.5 rounded text-xs font-semibold border ${RELATIONSHIP_STYLES[ownership.fields.RelationshipType]}`}>
                {ownership.fields.RelationshipType}
              </span>
            )}
            {ownership.fields.OwnershipPercent != null && (
              <span className="inline-block px-2 py-0.5 rounded text-xs font-bold bg-gray-100 text-gray-800 font-mono-data">
                {ownership.fields.OwnershipPercent}%
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500 mt-1">
            {[
              ownership.fields.ParentEntity && `Parent: ${ownership.fields.ParentEntity}`,
              linkedProperty && `Property: ${linkedProperty.fields.Title}`,
              ownership.fields.EffectiveDate && `Effective ${formatDate(ownership.fields.EffectiveDate)}`,
            ].filter(Boolean).join(' · ')}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              <button
                onClick={() => setShowDeleteConfirm(true)}
                className="px-3 py-1.5 border border-red-300 text-error hover:bg-red-50 rounded-md text-sm font-medium transition-colors"
              >
                Delete
              </button>
              <button
                onClick={handleEdit}
                className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors"
              >
                <Icon name="settings" size={14} />
                Edit
              </button>
            </>
          )}
          {editing && <EditingActionButtons saving={saving} onSave={handleSave} onCancel={handleCancel} />}
        </div>
      </div>

      <SaveErrorBanner error={saveError} />

      {showDeleteConfirm && (
        <div className="mb-4 bg-red-50 border border-red-200 rounded-lg p-4 flex items-start gap-3">
          <Icon name="alert" size={18} className="text-error flex-shrink-0 mt-0.5" />
          <div className="flex-1">
            <p className="text-sm font-semibold text-error mb-1">Delete this ownership record?</p>
            <p className="text-xs text-red-700 mb-3">
              This permanently removes the record from SharePoint. The Properties Registry and other lists are not affected.
            </p>
            <div className="flex gap-2">
              <button
                onClick={handleDelete}
                disabled={saving}
                className="px-3 py-1 bg-error hover:bg-red-700 text-white rounded text-sm font-medium disabled:opacity-50"
              >
                {saving ? 'Deleting…' : 'Yes, delete'}
              </button>
              <button
                onClick={() => setShowDeleteConfirm(false)}
                disabled={saving}
                className="px-3 py-1 border border-gray-300 text-gray-700 hover:bg-white rounded text-sm font-medium"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Section title="Entity">
          <EditableField
            label="Entity Name"
            value={display.Title}
            editing={editing}
            required
            onChange={(v) => handleFieldChange('Title', v as string)}
          />
          <EditableField
            label="Relationship Type"
            value={display.RelationshipType}
            editing={editing}
            type="choice"
            choices={CHOICES.RelationshipType}
            onChange={(v) => handleFieldChange('RelationshipType', v as RelationshipType)}
          />
          <EditableField
            label="Ownership %"
            value={display.OwnershipPercent}
            editing={editing}
            type="number"
            mono
            onChange={(v) => handleFieldChange('OwnershipPercent', v as number)}
          />
          <EditableField
            label="Member Class"
            value={display.MemberClass}
            editing={editing}
            type="choice"
            choices={['', 'Class A', 'Class B', 'Class C', 'Class D', 'N/A'] as const}
            onChange={(v) => handleFieldChange('MemberClass', (v as string) || undefined)}
          />
        </Section>

        <Section title="Hierarchy">
          <EditableField
            label="Parent Entity"
            value={display.ParentEntity}
            editing={editing}
            onChange={(v) => handleFieldChange('ParentEntity', v as string)}
          />
          {!editing && linkedProperty ? (
            <div className="flex items-start gap-3">
              <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Linked Property</dt>
              <dd className="text-sm flex-1">
                <button
                  onClick={() => navigate(`/properties/${linkedProperty.id}`)}
                  className="text-teal-700 hover:text-teal-900 font-medium underline"
                >
                  {linkedProperty.fields.Title}
                </button>
              </dd>
            </div>
          ) : editing ? (
            <div className="flex items-start gap-3">
              <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">Linked Property</label>
              <div className="flex-1">
                <select
                  value={String(display.LinkedPropertyLookupId ?? '')}
                  onChange={(e) => handleFieldChange('LinkedPropertyLookupId', e.target.value || undefined)}
                  className="w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 bg-white"
                >
                  <option value="">— none —</option>
                  {allProperties?.map((p) => (
                    <option key={p.id} value={p.id}>{p.fields.Title}</option>
                  ))}
                </select>
              </div>
            </div>
          ) : (
            <div className="flex items-start gap-3">
              <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Linked Property</dt>
              <dd className="text-sm text-gray-300 flex-1">—</dd>
            </div>
          )}
          <EditableField
            label="Effective Date"
            value={display.EffectiveDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('EffectiveDate', v as string)}
          />
        </Section>

        <Section title="Source">
          <EditableField
            label="Source Document"
            value={display.SourceDocument}
            editing={editing}
            onChange={(v) => handleFieldChange('SourceDocument', v as string)}
          />
        </Section>

        <Section title="Audit Trail">
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Created</dt>
            <dd className="text-sm text-gray-900 flex-1 font-mono-data text-xs">
              {formatDate(ownership.createdDateTime)}
            </dd>
          </div>
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Last Modified</dt>
            <dd className="text-sm text-gray-900 flex-1 font-mono-data text-xs">
              {formatDate(ownership.lastModifiedDateTime)}
            </dd>
          </div>
        </Section>

        <Section title="Notes" fullWidth>
          <EditableField
            label="Notes"
            value={display.EntityNotes}
            editing={editing}
            type="textarea"
            rows={5}
            hideLabel
            onChange={(v) => handleFieldChange('EntityNotes', v as string)}
          />
        </Section>
      </div>

      {/* PR-09d gap #5 — Ownership change reason modal */}
      {reasonModalOpen && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-lg shadow-xl max-w-md w-full p-5">
            <h3 className="text-lg font-bold text-teal-700 mb-2">Reason for change</h3>
            <p className="text-sm text-gray-600 mb-4">
              Per audit policy (Spec §5.5), every ownership change is logged with a reason. This will appear in the audit log alongside the diff.
            </p>
            <div className="space-y-3">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Reason <span className="text-error">*</span></label>
                <select
                  value={reason}
                  onChange={(e) => setReason(e.target.value as OwnershipChangeReason | '')}
                  className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white"
                  autoFocus
                >
                  <option value="">— select —</option>
                  {CHANGE_REASONS.map((r) => (
                    <option key={r} value={r}>{r}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Notes (optional)</label>
                <textarea
                  value={reasonNotes}
                  onChange={(e) => setReasonNotes(e.target.value)}
                  rows={3}
                  placeholder="Brief context, e.g., 'OA Amendment 2024, Stan exits to Cara'"
                  className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 resize-y"
                />
              </div>
            </div>
            <SaveErrorBanner error={saveError} />
            <div className="mt-5 flex items-center justify-end gap-2">
              <button
                onClick={() => {
                  setReasonModalOpen(false);
                  setPendingChanges(null);
                }}
                disabled={saving}
                className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
              >
                Cancel
              </button>
              <button
                onClick={confirmSave}
                disabled={saving || !reason}
                className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50"
              >
                {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
                {saving ? 'Saving…' : 'Save with Reason'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
