import { useState, useEffect, useMemo } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  PRIMARY_RELATIONSHIP_TYPES,
  type Property,
  type Owner,
  type RelationshipType,
  type OwnershipFields,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { toDateInputValue } from '../lib/dates';
import { BreadcrumbBar, Section, SaveErrorBanner } from '../components/detail';

export function OwnershipNew() {
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();

  const { data: allProperties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const { data: allOwners } = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  // Pre-fill from URL params (e.g., /ownership/new?parentOwnerId=5 from Owner Detail's Add Member)
  const prefilledParentOwnerId = searchParams.get('parentOwnerId');
  const prefilledPropertyId = searchParams.get('propertyId');

  const [form, setForm] = useState<Partial<OwnershipFields>>({
    OwnerLookupId: undefined,
    ParentOwnerLookupId: prefilledParentOwnerId ?? undefined,
    LinkedPropertyLookupId: prefilledPropertyId ?? undefined,
    RelationshipType: 'Member',
    OwnershipPercent: undefined,
    EffectiveDate: undefined,
    SourceDocument: '',
    EntityNotes: '',
  });

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Auto-derive Title from selected Owner's name (so SP Title column gets a sensible value)
  const selectedOwner = useMemo(() => {
    if (!form.OwnerLookupId || !allOwners) return null;
    return allOwners.find((o) => String(o.id) === String(form.OwnerLookupId)) ?? null;
  }, [form.OwnerLookupId, allOwners]);

  useEffect(() => {
    if (selectedOwner) {
      setForm((prev) => ({ ...prev, Title: selectedOwner.fields.Title }));
    }
  }, [selectedOwner?.id]);

  const handleChange = <K extends keyof OwnershipFields>(field: K, value: OwnershipFields[K]) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const handleSave = async () => {
    if (!form.OwnerLookupId) {
      setError('Owner is required — pick an existing owner or create a new one first.');
      return;
    }
    if (!form.LinkedPropertyLookupId && !form.ParentOwnerLookupId) {
      setError(
        'Either Linked Property or Parent Owner must be set — this row needs to point to something.'
      );
      return;
    }
    if (form.LinkedPropertyLookupId && form.ParentOwnerLookupId) {
      setError(
        'Set either Linked Property or Parent Owner, not both. One row = one relationship.'
      );
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const fields: Record<string, unknown> = {};
      Object.entries(form).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') fields[k] = v;
      });
      const created = await createListItem<{ id: string }>(LIST_NAMES.Ownership, fields);
      // Navigate back to the parent context
      if (form.ParentOwnerLookupId) {
        navigate(`/owners/${form.ParentOwnerLookupId}`);
      } else if (form.LinkedPropertyLookupId) {
        navigate(`/properties/${form.LinkedPropertyLookupId}`);
      } else {
        navigate(`/ownership/${created.id}`);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  const sortedOwners = useMemo(() => {
    if (!allOwners) return [];
    return [...allOwners].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [allOwners]);

  const sortedProperties = useMemo(() => {
    if (!allProperties) return [];
    return [...allProperties].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [allProperties]);

  return (
    <div>
      <BreadcrumbBar parentLabel="Ownership" parentTo="/ownership" currentLabel="New Ownership Entry" />

      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">New Ownership Entry</h1>
          <p className="text-sm text-gray-500 mt-1">
            Record an entity's ownership stake in a property or another entity.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => navigate(-1)}
            disabled={saving}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleSave}
            disabled={saving}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Saving…' : 'Create'}
          </button>
        </div>
      </div>

      <SaveErrorBanner error={error} />

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Section title="Who">
          <Field label="Owner" required>
            <select
              value={String(form.OwnerLookupId ?? '')}
              onChange={(e) => handleChange('OwnerLookupId', e.target.value || undefined)}
              className={`${inputClass} bg-white`}
              autoFocus
            >
              <option value="">— select an owner —</option>
              {sortedOwners.map((o) => (
                <option key={o.id} value={o.id}>
                  {o.fields.Title} {o.fields.OwnerType ? `(${o.fields.OwnerType})` : ''}
                </option>
              ))}
            </select>
            <p className="text-xs text-gray-400 mt-1">
              Owner not in the list? <button
                type="button"
                onClick={() => navigate('/owners/new')}
                className="text-teal-700 hover:text-teal-900 font-medium underline"
              >Create one first.</button>
            </p>
          </Field>

          <Field label="Relationship Type">
            <select
              value={form.RelationshipType ?? ''}
              onChange={(e) => handleChange('RelationshipType', (e.target.value || undefined) as RelationshipType)}
              className={`${inputClass} bg-white`}
            >
              <option value="">— select —</option>
              {PRIMARY_RELATIONSHIP_TYPES.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
            <p className="text-xs text-gray-500 mt-1 leading-relaxed">
              <strong>Managing Member</strong>: has manager authority per the OA (e.g., CAHP SC LLC at 0.01%). <br />
              <strong>Member</strong>: economic interest (capital + voting). Most owners are Members.
            </p>
          </Field>

          <Field label="Ownership %">
            <input
              type="number"
              value={form.OwnershipPercent ?? ''}
              onChange={(e) => handleChange('OwnershipPercent', e.target.value === '' ? undefined : Number(e.target.value))}
              placeholder="e.g., 50, 99.99, 0.01"
              className={`${inputClass} font-mono-data`}
              min={0}
              max={100}
              step="0.01"
            />
          </Field>
        </Section>

        <Section title="In What">
          <Field label="Linked Property">
            <select
              value={String(form.LinkedPropertyLookupId ?? '')}
              onChange={(e) =>
                handleChange('LinkedPropertyLookupId', e.target.value || undefined)
              }
              className={`${inputClass} bg-white`}
              disabled={!!form.ParentOwnerLookupId}
            >
              <option value="">— none —</option>
              {sortedProperties.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.fields.Title}
                </option>
              ))}
            </select>
          </Field>

          <Field label="OR Parent Owner">
            <select
              value={String(form.ParentOwnerLookupId ?? '')}
              onChange={(e) =>
                handleChange('ParentOwnerLookupId', e.target.value || undefined)
              }
              className={`${inputClass} bg-white`}
              disabled={!!form.LinkedPropertyLookupId}
            >
              <option value="">— none —</option>
              {sortedOwners.map((o) => (
                <option key={o.id} value={o.id}>
                  {o.fields.Title} {o.fields.OwnerType ? `(${o.fields.OwnerType})` : ''}
                </option>
              ))}
            </select>
            <p className="text-xs text-gray-400 mt-1">
              Use this when the owner is a member of another entity (e.g., "Stan owns 60% of VanRock Holdings").
            </p>
          </Field>

          <Field label="Effective Date">
            <input
              type="date"
              value={toDateInputValue(form.EffectiveDate)}
              onChange={(e) => handleChange('EffectiveDate', e.target.value ? new Date(e.target.value).toISOString() : undefined)}
              className={`${inputClass} font-mono-data`}
            />
          </Field>
        </Section>

        <Section title="Source & Notes" fullWidth>
          <Field label="Source Document">
            <input
              type="text"
              value={form.SourceDocument ?? ''}
              onChange={(e) => handleChange('SourceDocument', e.target.value)}
              placeholder="e.g., Operating Agreement v2.1 — Exhibit 1, Articles of Organization, Estate Settlement"
              className={inputClass}
            />
          </Field>
          <Field label="Notes">
            <textarea
              value={form.EntityNotes ?? ''}
              onChange={(e) => handleChange('EntityNotes', e.target.value)}
              placeholder="Class A/B distinction, voting rights, capital contributions, etc."
              rows={4}
              className={`${inputClass} resize-y`}
            />
          </Field>
        </Section>
      </div>

      <div className="mt-6 bg-blue-50 border border-blue-200 rounded-lg p-4 flex gap-3">
        <Icon name="alert" size={16} className="text-blue-700 flex-shrink-0 mt-0.5" />
        <div className="flex-1 text-xs text-blue-900">
          <p className="font-semibold mb-1">How to model ownership relationships</p>
          <p className="leading-relaxed">
            One row = one relationship. Set <strong>either</strong> Linked Property (this owner holds stake in a property) <strong>or</strong> Parent Owner (this owner is a member of another entity), not both.
            To trace beneficial ownership up through holding LLCs, create separate rows: e.g., row 1 = VanRock owns 99.99% of property X; row 2 = Stan owns 60% of VanRock.
          </p>
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, children }: { label: string; required?: boolean; children: React.ReactNode }) {
  return (
    <div className="flex items-start gap-3">
      <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
        {label}{required && <span className="text-error ml-0.5">*</span>}
      </label>
      <div className="flex-1">{children}</div>
    </div>
  );
}
