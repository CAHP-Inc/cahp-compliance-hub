import { useState, useEffect, useMemo } from 'react';
import {
  useSharePointList,
  useSharePointItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Ownership,
  type Owner,
  type Property,
  type RelationshipType,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';

interface EditOwnershipModalProps {
  ownershipId: string;
  onClose: () => void;
  onSaved: () => void;
}

const RELATIONSHIP_OPTIONS: RelationshipType[] = [
  'Managing Member',
  'Sole Member',
  'Member',
  'Owner',
  'Subsidiary',
  'Beneficial Owner',
];

const MEMBER_CLASS_OPTIONS = ['', 'Class A', 'Class B', 'Class C', 'Class D', 'N/A'];

const CHANGE_REASONS = ['Buy-In', 'Buy-Out', 'Estate', 'Initial Filing', 'Correction', 'Other'] as const;

/**
 * Full-record edit modal for an Ownership relationship.
 * Surfaces every editable field — Role, Class, %, Effective Date, Owner, Parent Owner,
 * Source Document, Notes — so future field additions automatically appear here without
 * UI changes.
 */
export function EditOwnershipModal({ ownershipId, onClose, onSaved }: EditOwnershipModalProps) {
  const ownership = useSharePointItem<Ownership>(LIST_NAMES.Ownership, ownershipId);
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [draft, setDraft] = useState<Ownership['fields'] | null>(null);
  const [reason, setReason] = useState<typeof CHANGE_REASONS[number]>('Correction');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (ownership.data && !draft) {
      setDraft({ ...ownership.data.fields });
    }
  }, [ownership.data, draft]);

  const sortedOwners = useMemo(() => {
    return [...(owners.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [owners.data]);

  const sortedProperties = useMemo(() => {
    return [...(properties.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')
    );
  }, [properties.data]);

  const update = <K extends keyof Ownership['fields']>(key: K, value: Ownership['fields'][K]) => {
    setDraft((d) => (d ? { ...d, [key]: value } : d));
  };

  const handleSave = async () => {
    if (!draft) return;
    setSaving(true);
    setError(null);
    try {
      const payload: Record<string, unknown> = {
        RelationshipType: draft.RelationshipType ?? null,
        OwnershipPercent: draft.OwnershipPercent ?? null,
        MemberClass: draft.MemberClass ?? null,
        EffectiveDate: draft.EffectiveDate ?? null,
        SourceDocument: draft.SourceDocument ?? null,
        EntityNotes: draft.EntityNotes ?? null,
        OwnerLookupId: draft.OwnerLookupId ? Number(draft.OwnerLookupId) : null,
        ParentOwnerLookupId: draft.ParentOwnerLookupId ? Number(draft.ParentOwnerLookupId) : null,
        LinkedPropertyLookupId: draft.LinkedPropertyLookupId
          ? Number(draft.LinkedPropertyLookupId)
          : null,
      };
      await updateListItem(LIST_NAMES.Ownership, ownershipId, payload, { reason });
      onSaved();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    const ok = window.confirm(
      'Delete this ownership record? The Owner entity will not be deleted — only this property-to-owner link. Action is logged.'
    );
    if (!ok) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Ownership, ownershipId);
      onSaved();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const isLoading = ownership.loading || owners.loading || properties.loading;
  const ownerName = draft?.OwnerLookupId
    ? sortedOwners.find((o) => String(o.id) === String(draft.OwnerLookupId))?.fields.Title
    : '(no owner)';

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !saving) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-3xl w-full max-h-[90vh] overflow-y-auto">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-200 flex items-start justify-between">
          <div>
            <h2 className="text-lg font-bold text-teal-700">Edit Ownership Record</h2>
            <p className="text-sm text-gray-500 mt-0.5">{ownerName ?? '(no owner)'}</p>
          </div>
          <button
            onClick={onClose}
            disabled={saving}
            className="text-gray-400 hover:text-gray-700 text-xl leading-none"
          >
            ×
          </button>
        </div>

        {isLoading || !draft ? (
          <div className="p-8 text-center text-sm text-gray-500">Loading…</div>
        ) : (
          <div className="px-6 py-4 space-y-4">
            {/* Owner */}
            <Row label="Owner">
              <select
                value={String(draft.OwnerLookupId ?? '')}
                onChange={(e) => update('OwnerLookupId', e.target.value || undefined)}
                disabled={saving}
                className={SELECT_CLASS}
              >
                <option value="">— Select an owner —</option>
                {sortedOwners.map((o) => (
                  <option key={o.id} value={String(o.id)}>
                    {o.fields.Title} {o.fields.OwnerType ? `· ${o.fields.OwnerType}` : ''}
                  </option>
                ))}
              </select>
            </Row>

            {/* Linked Property */}
            <Row label="Linked Property">
              <select
                value={String(draft.LinkedPropertyLookupId ?? '')}
                onChange={(e) => update('LinkedPropertyLookupId', e.target.value || undefined)}
                disabled={saving}
                className={SELECT_CLASS}
              >
                <option value="">— None (entity-level relationship) —</option>
                {sortedProperties.map((p) => (
                  <option key={p.id} value={String(p.id)}>
                    {p.fields.Title}
                  </option>
                ))}
              </select>
            </Row>

            {/* Parent Owner — for member-of-member chains */}
            <Row label="Parent Owner">
              <select
                value={String(draft.ParentOwnerLookupId ?? '')}
                onChange={(e) => update('ParentOwnerLookupId', e.target.value || undefined)}
                disabled={saving}
                className={SELECT_CLASS}
              >
                <option value="">— None —</option>
                {sortedOwners.map((o) => (
                  <option key={o.id} value={String(o.id)}>
                    {o.fields.Title}
                  </option>
                ))}
              </select>
              <p className="text-[11px] text-gray-500 mt-1">
                Set this when this entity is owned by another entity (e.g., CAHP SC LLC's parent is CAHP Inc.)
              </p>
            </Row>

            <div className="grid grid-cols-2 gap-4">
              {/* Relationship Type */}
              <Row label="Role / Relationship">
                <select
                  value={draft.RelationshipType ?? ''}
                  onChange={(e) => update('RelationshipType', (e.target.value as RelationshipType) || undefined)}
                  disabled={saving}
                  className={SELECT_CLASS}
                >
                  <option value="">— select —</option>
                  {RELATIONSHIP_OPTIONS.map((r) => (
                    <option key={r} value={r}>{r}</option>
                  ))}
                </select>
              </Row>

              {/* Member Class */}
              <Row label="Member Class">
                <select
                  value={draft.MemberClass ?? ''}
                  onChange={(e) => update('MemberClass', e.target.value || undefined)}
                  disabled={saving}
                  className={SELECT_CLASS}
                >
                  {MEMBER_CLASS_OPTIONS.map((c) => (
                    <option key={c} value={c}>{c || '—'}</option>
                  ))}
                </select>
              </Row>
            </div>

            <div className="grid grid-cols-2 gap-4">
              {/* Ownership Percent */}
              <Row label="Ownership %">
                <input
                  type="number"
                  step="0.01"
                  min={0}
                  max={100}
                  value={draft.OwnershipPercent ?? ''}
                  onChange={(e) => update('OwnershipPercent', e.target.value === '' ? undefined : Number(e.target.value))}
                  disabled={saving}
                  className={INPUT_CLASS}
                />
              </Row>

              {/* Effective Date */}
              <Row label="Effective Date">
                <input
                  type="date"
                  value={draft.EffectiveDate ? draft.EffectiveDate.slice(0, 10) : ''}
                  onChange={(e) => update('EffectiveDate', e.target.value || undefined)}
                  disabled={saving}
                  className={INPUT_CLASS}
                />
              </Row>
            </div>

            {/* Source Document */}
            <Row label="Source Document">
              <input
                type="text"
                value={draft.SourceDocument ?? ''}
                onChange={(e) => update('SourceDocument', e.target.value || undefined)}
                disabled={saving}
                placeholder="e.g., 'Amended OA dated 1/1/2024'"
                className={INPUT_CLASS}
              />
            </Row>

            {/* Notes */}
            <Row label="Notes">
              <textarea
                value={draft.EntityNotes ?? ''}
                onChange={(e) => update('EntityNotes', e.target.value || undefined)}
                disabled={saving}
                rows={3}
                className={`${INPUT_CLASS} resize-y`}
              />
            </Row>

            {/* Reason for change */}
            <Row label="Reason for Change">
              <select
                value={reason}
                onChange={(e) => setReason(e.target.value as typeof CHANGE_REASONS[number])}
                disabled={saving}
                className={SELECT_CLASS}
              >
                {CHANGE_REASONS.map((r) => (
                  <option key={r} value={r}>{r}</option>
                ))}
              </select>
              <p className="text-[11px] text-gray-500 mt-1">Captured in the audit log alongside the change.</p>
            </Row>

            {error && (
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
                <Icon name="alert" size={14} className="text-error flex-shrink-0 mt-0.5" />
                <p className="text-xs text-error font-mono-data">{error}</p>
              </div>
            )}
          </div>
        )}

        {/* Footer */}
        <div className="px-6 py-4 border-t border-gray-200 flex items-center justify-between gap-2 bg-gray-50 rounded-b-lg">
          <button
            onClick={handleDelete}
            disabled={saving || isLoading}
            className="text-xs text-error hover:text-red-700 font-medium px-3 py-1.5 rounded hover:bg-red-50 disabled:opacity-50"
          >
            Delete record
          </button>
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              disabled={saving}
              className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-100 rounded-md"
            >
              Cancel
            </button>
            <button
              onClick={handleSave}
              disabled={saving || isLoading || !draft}
              className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
            >
              {saving ? (
                <>
                  <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                  Saving…
                </>
              ) : (
                'Save'
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

const INPUT_CLASS =
  'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';
const SELECT_CLASS = INPUT_CLASS + ' bg-white';

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
        {label}
      </label>
      {children}
    </div>
  );
}
