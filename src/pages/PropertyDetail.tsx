import { useState, useEffect, useMemo } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type Property,
  type PropertyFields,
  type Submittal,
  type PropertyStatus,
  type SubmittalStatusValue,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';

const STATUS_STYLES: Record<PropertyStatus, string> = {
  Active: 'bg-green-100 text-green-800 border-green-200',
  Pending: 'bg-yellow-100 text-yellow-800 border-yellow-200',
  Withdrawn: 'bg-gray-100 text-gray-700 border-gray-200',
  'Removed from Program': 'bg-red-100 text-red-800 border-red-200',
  Sold: 'bg-blue-100 text-blue-800 border-blue-200',
};

const SUBMITTAL_STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  Draft: 'bg-gray-100 text-gray-700',
  'Package Mailed (NC)': 'bg-blue-100 text-blue-800',
  Filed: 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-yellow-100 text-yellow-800',
  'Responded - Awaiting DOR': 'bg-amber-100 text-amber-800',
  Approved: 'bg-green-100 text-green-800',
  Denied: 'bg-red-100 text-red-800',
  Withdrawn: 'bg-gray-100 text-gray-700',
};

// Choice options matching SharePoint config exactly
const CHOICES = {
  AMIProgram: ['20/50', '40/60', 'Mixed', 'None'] as const,
  CAHPLanguageAdded: ['Yes', 'No', 'In Progress', 'Needs Revision'] as const,
  cahpCounty: [
    'Greenville (SC)', 'Spartanburg (SC)', 'Anderson (SC)', 'Pickens (SC)',
    'Laurens (SC)', 'York (SC)', 'Mecklenburg (NC)', 'Guilford (NC)',
    'Durham (NC)', 'Wake (NC)', 'Forsyth (NC)', 'Buncombe (NC)', 'Other',
  ] as const,
  cahpOwnerGroup: [
    'VanRock Holdings', 'Red Cedar', 'AmRock', 'Troy Hampton',
    'Deepak', 'Damon Lilly', 'Other',
  ] as const,
  cahpState: ['SC', 'NC'] as const,
  cahpVerificationStatus: [
    'Inherited - Unverified', 'Verified', 'Needs Follow-Up', 'N/A',
  ] as const,
  LURAExecuted: ['Yes', 'No', 'In Progress', 'N/A'] as const,
  PropertyStatus: ['Active', 'Pending', 'Withdrawn', 'Removed from Program', 'Sold'] as const,
} as const;

type TabId = 'overview' | 'submittals' | 'documents' | 'compliance' | 'notes';

const TABS: { id: TabId; label: string; shipped: boolean }[] = [
  { id: 'overview', label: 'Overview', shipped: true },
  { id: 'submittals', label: 'Submittals', shipped: true },
  { id: 'documents', label: 'Documents', shipped: false },
  { id: 'compliance', label: 'Compliance', shipped: false },
  { id: 'notes', label: 'Notes', shipped: false },
];

export function PropertyDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [activeTab, setActiveTab] = useState<TabId>('overview');

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<PropertyFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  const { data: property, loading, error, refetch } = useSharePointItem<Property>(
    LIST_NAMES.Properties, id
  );

  const { data: allSubmittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, {
    top: 500,
  });

  const relatedSubmittals = useMemo(() => {
    if (!allSubmittals || !id) return [];
    return allSubmittals.filter((s) => String(s.fields.PropertyLookupId) === String(id));
  }, [allSubmittals, id]);

  // Reset draft when underlying property changes (and we're not in the middle of editing)
  useEffect(() => {
    if (property && !editing) {
      setDraft({ ...property.fields });
    }
  }, [property?.id, property?.lastModifiedDateTime, editing]);

  const handleEdit = () => {
    if (!property) return;
    setDraft({ ...property.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    if (property) setDraft({ ...property.fields });
    setSaveError(null);
    setEditing(false);
  };

  const handleFieldChange = <K extends keyof PropertyFields>(
    field: K, value: PropertyFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleSave = async () => {
    if (!draft || !property) return;
    setSaving(true);
    setSaveError(null);
    try {
      const changes: Record<string, unknown> = {};
      (Object.keys(draft) as (keyof PropertyFields)[]).forEach((k) => {
        const oldVal = property.fields[k];
        const newVal = draft[k];
        if (oldVal !== newVal) {
          changes[k as string] = newVal === '' ? null : newVal;
        }
      });
      if (Object.keys(changes).length > 0) {
        await updateListItem(LIST_NAMES.Properties, property.id, changes);
        await refetch();
      }
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return (
      <div className="text-center py-12 text-gray-500">
        <div className="inline-flex items-center gap-3">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading property…</span>
        </div>
      </div>
    );
  }

  if (error || !property || !draft) {
    return (
      <div>
        <BreadcrumbBar />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Property not found
          </div>
          <p className="text-sm text-red-700 mb-3">
            {error ? error.message : `No property with ID ${id} exists in Properties Registry.`}
          </p>
          <button
            onClick={() => navigate('/properties')}
            className="text-sm text-teal-700 hover:text-teal-900 font-medium underline"
          >
            ← Back to Properties
          </button>
        </div>
      </div>
    );
  }

  const display = editing ? draft : property.fields;

  return (
    <div>
      <BreadcrumbBar propertyName={property.fields.Title} />

      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-3 flex-wrap">
            <h1 className="text-3xl font-bold text-teal-700">{property.fields.Title}</h1>
            {property.fields.PropertyStatus && (
              <span className={`inline-block px-2.5 py-0.5 rounded text-xs font-semibold border ${STATUS_STYLES[property.fields.PropertyStatus]}`}>
                {property.fields.PropertyStatus}
              </span>
            )}
            {property.fields.cahpVerificationStatus === 'Inherited - Unverified' && (
              <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-yellow-50 text-yellow-800 border border-yellow-200">
                UNVERIFIED
              </span>
            )}
            {property.fields.cahpVerificationStatus === 'Needs Follow-Up' && (
              <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-error/10 text-error border border-error/20">
                NEEDS FOLLOW-UP
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500 mt-1">
            {[
              property.fields.LegalEntity,
              property.fields.cahpState,
              property.fields.cahpCounty?.replace(/\s*\([^)]*\)\s*/g, ''),
              property.fields.UnitCount ? `${property.fields.UnitCount} units` : null,
              property.fields.AMIProgram && property.fields.AMIProgram !== 'None'
                ? `${property.fields.AMIProgram} AMI` : null,
              property.fields.cahpOwnerGroup,
            ].filter(Boolean).join(' · ')}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              <button
                className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
                disabled title="Watch button ships in a future PR"
              >
                <Icon name="star" size={14} />
                Watch
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
          {editing && (
            <>
              <button
                onClick={handleCancel} disabled={saving}
                className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
              >
                Cancel
              </button>
              <button
                onClick={handleSave} disabled={saving}
                className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
                {saving ? 'Saving…' : 'Save Changes'}
              </button>
            </>
          )}
        </div>
      </div>

      {saveError && (
        <div className="mb-4 bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
          <Icon name="alert" size={16} className="text-error flex-shrink-0 mt-0.5" />
          <div className="flex-1 text-sm">
            <span className="font-semibold text-error">Save failed.</span>{' '}
            <span className="text-red-700 font-mono-data text-xs">{saveError}</span>
          </div>
        </div>
      )}

      <div className="border-b border-gray-200 mb-6">
        <div className="flex gap-1">
          {TABS.map((t) => {
            const active = activeTab === t.id;
            const disabled = !t.shipped || editing;
            return (
              <button
                key={t.id}
                onClick={() => !disabled && setActiveTab(t.id)}
                disabled={disabled}
                className={`px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
                  active ? 'border-teal-700 text-teal-700'
                    : !disabled ? 'border-transparent text-gray-600 hover:text-teal-700 hover:border-gray-300'
                    : 'border-transparent text-gray-400 cursor-not-allowed'
                }`}
                title={editing ? 'Save or cancel to switch tabs' : !t.shipped ? 'Coming in a future PR' : undefined}
              >
                {t.label}
                {!t.shipped && <span className="ml-1.5 text-[10px] text-gray-400">soon</span>}
              </button>
            );
          })}
        </div>
      </div>

      {activeTab === 'overview' && (
        <OverviewTab display={display} editing={editing} onChange={handleFieldChange} />
      )}
      {activeTab === 'submittals' && <SubmittalsTab submittals={relatedSubmittals} />}
    </div>
  );
}

function OverviewTab({
  display, editing, onChange,
}: {
  display: PropertyFields;
  editing: boolean;
  onChange: <K extends keyof PropertyFields>(field: K, value: PropertyFields[K]) => void;
}) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <Section title="Property Identity">
        <EditableField label="Property Name" value={display.Title} editing={editing} required onChange={(v) => onChange('Title', v as string)} />
        <EditableField label="Legal Entity" value={display.LegalEntity} editing={editing} onChange={(v) => onChange('LegalEntity', v as string)} />
        <EditableField label="Property Address" value={display.PropertyAddress} editing={editing} onChange={(v) => onChange('PropertyAddress', v as string)} />
        <EditableField label="DOR Account ID" value={display.DORAccountID} editing={editing} mono onChange={(v) => onChange('DORAccountID', v as string)} />
        <EditableField label="Date Added to CAHP" value={display.DateAddedToCAHP} editing={editing} type="date" onChange={(v) => onChange('DateAddedToCAHP', v as string)} />
      </Section>

      <Section title="Affordability">
        <EditableField label="AMI Program" value={display.AMIProgram} editing={editing} type="choice" choices={CHOICES.AMIProgram} onChange={(v) => onChange('AMIProgram', v as PropertyFields['AMIProgram'])} />
        <EditableField label="CAHP Language Added" value={display.CAHPLanguageAdded} editing={editing} type="choice" choices={CHOICES.CAHPLanguageAdded} onChange={(v) => onChange('CAHPLanguageAdded', v as PropertyFields['CAHPLanguageAdded'])} />
        <EditableField label="LURA Executed" value={display.LURAExecuted} editing={editing} type="choice" choices={CHOICES.LURAExecuted} onChange={(v) => onChange('LURAExecuted', v as PropertyFields['LURAExecuted'])} />
        <EditableField label="Operating Agreement Version" value={display.OpAgreementVersion} editing={editing} mono onChange={(v) => onChange('OpAgreementVersion', v as string)} />
      </Section>

      <Section title="Status & Verification">
        <EditableField label="Property Status" value={display.PropertyStatus} editing={editing} type="choice" choices={CHOICES.PropertyStatus} onChange={(v) => onChange('PropertyStatus', v as PropertyFields['PropertyStatus'])} />
        <EditableField label="Verification Status" value={display.cahpVerificationStatus} editing={editing} type="choice" choices={CHOICES.cahpVerificationStatus} onChange={(v) => onChange('cahpVerificationStatus', v as PropertyFields['cahpVerificationStatus'])} />
        <EditableField label="Removed Reason" value={display.RemovedReason} editing={editing} type="textarea" onChange={(v) => onChange('RemovedReason', v as string)} />
      </Section>

      <Section title="Location & Ownership">
        <EditableField label="State" value={display.cahpState} editing={editing} type="choice" choices={CHOICES.cahpState} mono onChange={(v) => onChange('cahpState', v as PropertyFields['cahpState'])} />
        <EditableField label="County" value={display.cahpCounty} editing={editing} type="choice" choices={CHOICES.cahpCounty} onChange={(v) => onChange('cahpCounty', v as string)} />
        <EditableField label="Units" value={display.UnitCount} editing={editing} type="number" mono onChange={(v) => onChange('UnitCount', v as number)} />
        <EditableField label="Owner Group" value={display.cahpOwnerGroup} editing={editing} type="choice" choices={CHOICES.cahpOwnerGroup} onChange={(v) => onChange('cahpOwnerGroup', v as PropertyFields['cahpOwnerGroup'])} />
      </Section>

      <Section title="Notes" fullWidth>
        <EditableField
          label="Property Notes" value={display.PropertyNotes} editing={editing}
          type="textarea" rows={5} hideLabel
          onChange={(v) => onChange('PropertyNotes', v as string)}
        />
      </Section>
    </div>
  );
}

function SubmittalsTab({ submittals }: { submittals: Submittal[] }) {
  if (submittals.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
        <p className="text-sm text-gray-500">No submittals on file for this property yet.</p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <table className="w-full text-sm">
        <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
          <tr>
            <th className="px-4 py-3 text-left">Submittal</th>
            <th className="px-4 py-3 text-left">Tax Year</th>
            <th className="px-4 py-3 text-left">Filing Method</th>
            <th className="px-4 py-3 text-left">Status</th>
            <th className="px-4 py-3 text-left">Filed Date</th>
            <th className="px-4 py-3 text-right">Approved Abatement</th>
            <th className="px-4 py-3 text-left">Next Action</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {submittals.map((s) => (
            <tr key={s.id} className="hover:bg-gray-50 transition-colors">
              <td className="px-4 py-3 font-medium text-gray-900">{s.fields.Title}</td>
              <td className="px-4 py-3 text-gray-700 font-mono-data">{s.fields.cahpTaxYear || '—'}</td>
              <td className="px-4 py-3 text-gray-700 text-xs">{s.fields.FilingMethod || '—'}</td>
              <td className="px-4 py-3">
                {s.fields.SubmittalStatus ? (
                  <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${SUBMITTAL_STATUS_STYLES[s.fields.SubmittalStatus] || 'bg-gray-100'}`}>
                    {s.fields.SubmittalStatus}
                  </span>
                ) : '—'}
              </td>
              <td className="px-4 py-3 text-gray-700 text-xs">{formatDate(s.fields.DateFiled)}</td>
              <td className="px-4 py-3 text-right font-mono-data">
                {s.fields.ApprovedAbatement ? `$${s.fields.ApprovedAbatement.toLocaleString()}` : '—'}
              </td>
              <td className="px-4 py-3 text-gray-700 text-xs">{s.fields.NextAction || '—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function BreadcrumbBar({ propertyName }: { propertyName?: string }) {
  return (
    <nav className="mb-4 text-sm">
      <Link to="/properties" className="text-teal-700 hover:text-teal-900 font-medium">
        ← Properties
      </Link>
      {propertyName && (
        <>
          <span className="mx-2 text-gray-400">/</span>
          <span className="text-gray-700">{propertyName}</span>
        </>
      )}
    </nav>
  );
}

function Section({
  title, children, fullWidth,
}: {
  title: string;
  children: React.ReactNode;
  fullWidth?: boolean;
}) {
  return (
    <div className={`bg-white border border-gray-200 rounded-lg p-5 shadow-card ${fullWidth ? 'lg:col-span-2' : ''}`}>
      <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">{title}</h3>
      <div className="space-y-2.5">{children}</div>
    </div>
  );
}

type FieldValue = string | number | null | undefined;
type FieldType = 'text' | 'textarea' | 'number' | 'date' | 'choice';

function EditableField({
  label, value, editing, onChange,
  type = 'text', choices, mono, required, rows = 3, hideLabel,
}: {
  label: string;
  value: FieldValue;
  editing: boolean;
  onChange: (v: FieldValue) => void;
  type?: FieldType;
  choices?: readonly string[];
  mono?: boolean;
  required?: boolean;
  rows?: number;
  hideLabel?: boolean;
}) {
  if (!editing) {
    let displayValue: React.ReactNode = value;
    if (type === 'date' && value) displayValue = formatDate(value as string);
    if (type === 'textarea' && value) {
      displayValue = <span className="whitespace-pre-wrap leading-relaxed block">{String(value)}</span>;
    }
    return (
      <div className="flex items-start gap-3">
        {!hideLabel && <dt className="text-sm text-gray-500 w-44 flex-shrink-0">{label}</dt>}
        <dd className={`text-sm text-gray-900 flex-1 ${mono ? 'font-mono-data' : ''}`}>
          {value === null || value === undefined || value === '' ? (
            <span className="text-gray-300">—</span>
          ) : displayValue}
        </dd>
      </div>
    );
  }

  const inputClass = 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  let input: React.ReactNode;
  if (type === 'textarea') {
    input = (
      <textarea value={String(value ?? '')} onChange={(e) => onChange(e.target.value)} rows={rows} className={`${inputClass} resize-y`} />
    );
  } else if (type === 'number') {
    input = (
      <input type="number" value={value == null || value === '' ? '' : value}
        onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}
        className={`${inputClass} ${mono ? 'font-mono-data' : ''}`} />
    );
  } else if (type === 'date') {
    const dateValue = value ? new Date(value as string).toISOString().slice(0, 10) : '';
    input = (
      <input type="date" value={dateValue}
        onChange={(e) => onChange(e.target.value ? new Date(e.target.value).toISOString() : null)}
        className={`${inputClass} font-mono-data`} />
    );
  } else if (type === 'choice') {
    input = (
      <select value={String(value ?? '')} onChange={(e) => onChange(e.target.value || null)}
        className={`${inputClass} bg-white`}>
        <option value="">— none —</option>
        {choices?.map((c) => (<option key={c} value={c}>{c}</option>))}
      </select>
    );
  } else {
    input = (
      <input type="text" value={String(value ?? '')} onChange={(e) => onChange(e.target.value)}
        className={`${inputClass} ${mono ? 'font-mono-data' : ''}`} required={required} />
    );
  }

  return (
    <div className="flex items-start gap-3">
      {!hideLabel && (
        <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
          {label}{required && <span className="text-error ml-0.5">*</span>}
        </label>
      )}
      <div className="flex-1">{input}</div>
    </div>
  );
}

function formatDate(iso: string | undefined | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
