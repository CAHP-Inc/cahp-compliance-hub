import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Property,
  type RelationshipType,
  type OwnershipFields,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { BreadcrumbBar, Section, SaveErrorBanner } from '../components/detail';

const CHOICES = {
  RelationshipType: [
    'Managing Member', 'Member', 'Owner', 'Subsidiary', 'Beneficial Owner',
  ] as const,
} as const;

export function OwnershipNew() {
  const navigate = useNavigate();
  const { data: allProperties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 200 });

  const [form, setForm] = useState<Partial<OwnershipFields>>({
    Title: '',
    RelationshipType: undefined,
    OwnershipPercent: undefined,
    ParentEntity: '',
    LinkedPropertyLookupId: undefined,
    EffectiveDate: undefined,
    SourceDocument: '',
    EntityNotes: '',
  });

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleChange = <K extends keyof OwnershipFields>(field: K, value: OwnershipFields[K]) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const handleSave = async () => {
    if (!form.Title || form.Title.trim() === '') {
      setError('Entity Name is required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // Drop empty/undefined values so SharePoint doesn't store them
      const fields: Record<string, unknown> = {};
      Object.entries(form).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') fields[k] = v;
      });
      const created = await createListItem<{ id: string }>(LIST_NAMES.Ownership, fields);
      navigate(`/ownership/${created.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  return (
    <div>
      <BreadcrumbBar parentLabel="Ownership" parentTo="/ownership" currentLabel="New Entry" />

      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">New Ownership Entry</h1>
          <p className="text-sm text-gray-500 mt-1">
            Record an entity's relationship to a property or parent entity.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => navigate('/ownership')}
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
        <Section title="Entity">
          <FormField label="Entity Name" required>
            <input
              type="text"
              value={form.Title ?? ''}
              onChange={(e) => handleChange('Title', e.target.value)}
              placeholder="e.g., VanRock Holdings, Stan Gendlin, 144 W Henry LLC"
              className={inputClass}
              autoFocus
            />
          </FormField>

          <FormField label="Relationship Type">
            <select
              value={form.RelationshipType ?? ''}
              onChange={(e) => handleChange('RelationshipType', (e.target.value || undefined) as RelationshipType)}
              className={`${inputClass} bg-white`}
            >
              <option value="">— select —</option>
              {CHOICES.RelationshipType.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </FormField>

          <FormField label="Ownership %">
            <input
              type="number"
              value={form.OwnershipPercent ?? ''}
              onChange={(e) => handleChange('OwnershipPercent', e.target.value === '' ? undefined : Number(e.target.value))}
              placeholder="e.g., 51"
              className={`${inputClass} font-mono-data`}
              min={0}
              max={100}
              step="0.01"
            />
          </FormField>
        </Section>

        <Section title="Hierarchy">
          <FormField label="Parent Entity">
            <input
              type="text"
              value={form.ParentEntity ?? ''}
              onChange={(e) => handleChange('ParentEntity', e.target.value)}
              placeholder="e.g., VanRock Holdings"
              className={inputClass}
            />
          </FormField>

          <FormField label="Linked Property">
            <select
              value={String(form.LinkedPropertyLookupId ?? '')}
              onChange={(e) => handleChange('LinkedPropertyLookupId', e.target.value || undefined)}
              className={`${inputClass} bg-white`}
            >
              <option value="">— none —</option>
              {allProperties?.map((p) => (
                <option key={p.id} value={p.id}>{p.fields.Title}</option>
              ))}
            </select>
          </FormField>

          <FormField label="Effective Date">
            <input
              type="date"
              value={form.EffectiveDate ? new Date(form.EffectiveDate).toISOString().slice(0, 10) : ''}
              onChange={(e) => handleChange('EffectiveDate', e.target.value ? new Date(e.target.value).toISOString() : undefined)}
              className={`${inputClass} font-mono-data`}
            />
          </FormField>
        </Section>

        <Section title="Source" fullWidth>
          <FormField label="Source Document">
            <input
              type="text"
              value={form.SourceDocument ?? ''}
              onChange={(e) => handleChange('SourceDocument', e.target.value)}
              placeholder="e.g., Operating Agreement v2.1, Articles of Organization, Estate Settlement"
              className={inputClass}
            />
          </FormField>
        </Section>

        <Section title="Notes" fullWidth>
          <textarea
            value={form.EntityNotes ?? ''}
            onChange={(e) => handleChange('EntityNotes', e.target.value)}
            placeholder="Any additional context — beneficiary structure, voting rights, restrictions, etc."
            rows={5}
            className={`${inputClass} resize-y`}
          />
        </Section>
      </div>

      <div className="mt-6 bg-blue-50 border border-blue-200 rounded-lg p-4 flex gap-3">
        <Icon name="alert" size={16} className="text-blue-700 flex-shrink-0 mt-0.5" />
        <div className="flex-1 text-xs text-blue-900">
          <p className="font-semibold mb-1">Tip on modeling ownership chains</p>
          <p className="leading-relaxed">
            For "X is the Managing Member of property LLC Y" → Entity Name = X, Relationship = Managing Member, Linked Property = Y.
            For "X owns 100% of holding entity Z" → Entity Name = X, Relationship = Owner, Parent Entity = Z (no Linked Property).
            Stacking entries this way lets you trace beneficial ownership up through holding LLCs to natural-person owners.
          </p>
        </div>
      </div>
    </div>
  );
}

function FormField({ label, required, children }: { label: string; required?: boolean; children: React.ReactNode }) {
  return (
    <div className="flex items-start gap-3">
      <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
        {label}{required && <span className="text-error ml-0.5">*</span>}
      </label>
      <div className="flex-1">{children}</div>
    </div>
  );
}
