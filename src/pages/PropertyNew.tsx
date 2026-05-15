import { useState, useMemo } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import {
  createListItem,
  useSharePointList,
  LIST_NAMES,
  type PropertyFields,
  type AMIProgram,
  type CAHPLanguageStatus,
  type CahpState,
  type CahpTaxYear,
  type LURAExecutedStatus,
  type OwnerGroup,
  type PropertyStatus,
  type VerificationStatus,
  type Owner,
  type SubmittalStatusValue,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { BreadcrumbBar, Section, SaveErrorBanner } from '../components/detail';

const CHOICES = {
  AMIProgram: ['20/50', '40/60', '50/80', '60/80', 'Mixed', 'None'] as const,
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
  TaxYear: ['2023', '2024', '2025', '2026', '2027', '2028'] as const,
  FilingType: ['Initial', 'Annual'] as const,
} as const;

type WizardStep = 1 | 2 | 3 | 4 | 5 | 6;

const STEP_TITLES: Record<WizardStep, { title: string; subtitle: string }> = {
  1: { title: 'Identity & Location', subtitle: 'Who is this property and where is it?' },
  2: { title: 'Property Details', subtitle: 'Units, AMI program, lifecycle.' },
  3: { title: 'Ownership', subtitle: 'Direct members of this property LLC.' },
  4: { title: 'Filing Config', subtitle: 'CAHP fee and first submittal setup.' },
  5: { title: 'Initial Documents', subtitle: 'Auto-created Outstanding Items for document collection.' },
  6: { title: 'Review & Create', subtitle: 'Last look before writing everything to SharePoint.' },
};

// The 7 initial-documents items per Spec §3.3.3 Step 5
const INITIAL_DOCS = [
  { title: 'Operating Agreement (executed)', category: 'Operating Agreement' as const },
  { title: 'LURA executed', category: 'LURA' as const },
  { title: 'Current AMI Certification', category: 'AMI Certification' as const },
  { title: 'Organizational Chart', category: 'Org Chart' as const },
  { title: 'Current Rent Roll', category: 'Rent Roll' as const },
  { title: 'Property Deed', category: 'Deed' as const },
  { title: 'Insurance Certificate', category: 'Other' as const },
];

interface OwnerEntryDraft {
  ownerLookupId: string;
  role: 'Managing Member' | 'Member';
  percent: number;
}

export function PropertyNew() {
  const navigate = useNavigate();
  const [step, setStep] = useState<WizardStep>(1);

  const { data: allOwners } = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const [form, setForm] = useState<Partial<PropertyFields>>({
    Title: '',
    LegalEntity: '',
    PropertyAddress: '',
    cahpState: undefined,
    cahpCounty: undefined,
    cahpOwnerGroup: undefined,
    DateAddedToCAHP: new Date().toISOString(),
    AMIProgram: undefined,
    CAHPLanguageAdded: undefined,
    LURAExecuted: undefined,
    OpAgreementVersion: '',
    UnitCount: undefined,
    DORAccountID: '',
    PropertyStatus: 'Pending',
    cahpVerificationStatus: 'Inherited - Unverified',
    PropertyNotes: '',
  });

  const [filingConfig, setFilingConfig] = useState({
    cahpFeePercent: 50, // typical CAHP fee share
    firstFilingType: 'Initial' as 'Initial' | 'Annual',
    firstFilingTaxYear: new Date().getFullYear().toString() as CahpTaxYear,
    annualFilingRequired: true,
  });

  // Owner entries — CAHP SC LLC pre-added at 0.01% if it exists as an Owner
  const cahpScLLC = useMemo(() => {
    if (!allOwners) return null;
    return (
      allOwners.find((o) => o.fields.Title === 'CAHP SC LLC') ??
      allOwners.find((o) =>
        (o.fields.Title ?? '').toLowerCase().includes('cahp sc')
      ) ??
      null
    );
  }, [allOwners]);

  const [ownerEntries, setOwnerEntries] = useState<OwnerEntryDraft[]>([]);

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);

  const handleChange = <K extends keyof PropertyFields>(field: K, value: PropertyFields[K]) => {
    setForm((prev) => ({ ...prev, [field]: value }));
    setValidationError(null);
  };

  const validateStep = (s: WizardStep): string | null => {
    if (s === 1) {
      if (!form.Title || form.Title.trim() === '') return 'Property Name is required.';
      if (!form.cahpState) return 'State (SC or NC) is required.';
    }
    return null;
  };

  const goNext = () => {
    const err = validateStep(step);
    if (err) { setValidationError(err); return; }
    if (step < 6) setStep((step + 1) as WizardStep);
  };

  const goBack = () => {
    setValidationError(null);
    if (step > 1) setStep((step - 1) as WizardStep);
  };

  const handleCreate = async () => {
    const err = validateStep(1);
    if (err) { setValidationError(err); setStep(1); return; }

    setSaving(true);
    setError(null);
    try {
      // 1. Create the Property record
      const propertyFields: Record<string, unknown> = {};
      Object.entries(form).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') propertyFields[k] = v;
      });
      const property = await createListItem<{ id: string }>(LIST_NAMES.Properties, propertyFields);
      const propertyId = property.id;

      // 2. Create the first Submittal
      try {
        const submittalTitle = `${form.Title} — ${filingConfig.firstFilingType} ${filingConfig.firstFilingTaxYear}`;
        await createListItem(LIST_NAMES.Submittals, {
          Title: submittalTitle,
          PropertyLookupId: propertyId,
          cahpTaxYear: filingConfig.firstFilingTaxYear,
          cahpState: form.cahpState,
          SubmittalStatus: 'Draft' as SubmittalStatusValue,
          NextAction: 'Gather initial documents',
        });
      } catch (e) {
        // eslint-disable-next-line no-console
        console.warn('First Submittal creation failed:', e);
      }

      // 3. Create the 7 Outstanding Items
      const today = new Date().toISOString();
      for (const doc of INITIAL_DOCS) {
        try {
          await createListItem(LIST_NAMES.Outstanding, {
            Title: doc.title,
            PropertyLookupId: propertyId,
            ItemCategory: doc.category,
            ItemStatus: 'Requested',
            DateRequested: today,
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn(`Outstanding Item "${doc.title}" creation failed:`, e);
        }
      }

      // 4. Create Ownership records
      // Auto-add CAHP SC LLC at 0.01% Managing Member if it exists
      if (cahpScLLC) {
        try {
          await createListItem(LIST_NAMES.Ownership, {
            Title: cahpScLLC.fields.Title,
            OwnerLookupId: cahpScLLC.id,
            LinkedPropertyLookupId: propertyId,
            RelationshipType: 'Managing Member',
            OwnershipPercent: 0.01,
            SourceDocument: 'Property Creation Wizard',
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn('CAHP SC LLC ownership row creation failed:', e);
        }
      }

      // Plus the user-added members
      for (const entry of ownerEntries) {
        const owner = allOwners?.find((o) => String(o.id) === String(entry.ownerLookupId));
        if (!owner) continue;
        try {
          await createListItem(LIST_NAMES.Ownership, {
            Title: owner.fields.Title,
            OwnerLookupId: entry.ownerLookupId,
            LinkedPropertyLookupId: propertyId,
            RelationshipType: entry.role,
            OwnershipPercent: entry.percent,
            SourceDocument: 'Property Creation Wizard',
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn(`Ownership row for ${owner.fields.Title} failed:`, e);
        }
      }

      navigate(`/properties/${propertyId}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  const addOwnerEntry = () => {
    setOwnerEntries((prev) => [
      ...prev,
      { ownerLookupId: '', role: 'Member', percent: 0 },
    ]);
  };

  const updateOwnerEntry = (idx: number, patch: Partial<OwnerEntryDraft>) => {
    setOwnerEntries((prev) => prev.map((e, i) => (i === idx ? { ...e, ...patch } : e)));
  };

  const removeOwnerEntry = (idx: number) => {
    setOwnerEntries((prev) => prev.filter((_, i) => i !== idx));
  };

  const totalOwnerPercent =
    (cahpScLLC ? 0.01 : 0) + ownerEntries.reduce((sum, e) => sum + (e.percent || 0), 0);

  return (
    <div>
      <BreadcrumbBar parentLabel="Properties" parentTo="/properties" currentLabel="New Property" />

      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">New Property</h1>
        <p className="text-sm text-gray-500 mt-1">{STEP_TITLES[step].subtitle}</p>
      </div>

      {/* Stepper */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-3">
        <div className="flex items-center gap-1 sm:gap-2 overflow-x-auto">
          {([1, 2, 3, 4, 5, 6] as WizardStep[]).map((s, idx) => {
            const isCurrent = s === step;
            const isComplete = s < step;
            return (
              <div key={s} className="flex items-center gap-1 sm:gap-2 flex-1 min-w-0">
                <div className="flex items-center gap-2 min-w-0">
                  <div
                    className={`flex-shrink-0 w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold ${
                      isCurrent
                        ? 'bg-teal-700 text-white'
                        : isComplete
                          ? 'bg-success text-white'
                          : 'bg-gray-100 text-gray-400'
                    }`}
                  >
                    {isComplete ? <Icon name="check" size={12} /> : s}
                  </div>
                  <div className="hidden md:block min-w-0">
                    <div className={`text-[10px] font-semibold uppercase tracking-wider ${isCurrent ? 'text-teal-700' : isComplete ? 'text-success' : 'text-gray-400'}`}>
                      Step {s}
                    </div>
                    <div className={`text-xs truncate ${isCurrent ? 'text-gray-900 font-medium' : 'text-gray-500'}`}>
                      {STEP_TITLES[s].title}
                    </div>
                  </div>
                </div>
                {idx < 5 && <div className={`flex-1 h-0.5 min-w-[20px] ${isComplete ? 'bg-success' : 'bg-gray-200'}`} />}
              </div>
            );
          })}
        </div>
      </div>

      <SaveErrorBanner error={error} />
      {validationError && (
        <div className="mb-4 bg-yellow-50 border border-yellow-200 rounded-lg p-3 flex items-start gap-2">
          <Icon name="alert" size={16} className="text-yellow-700 flex-shrink-0 mt-0.5" />
          <p className="text-sm text-yellow-800">{validationError}</p>
        </div>
      )}

      {/* Step 1 — Identity & Location */}
      {step === 1 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Section title="Identity">
            <Field label="Property Name" required>
              <input type="text" value={form.Title ?? ''} onChange={(e) => handleChange('Title', e.target.value)} placeholder="e.g., 135 Oakwood Apartments" className={inputClass} autoFocus />
            </Field>
            <Field label="Legal Entity">
              <input type="text" value={form.LegalEntity ?? ''} onChange={(e) => handleChange('LegalEntity', e.target.value)} placeholder="e.g., 135 Oakwood LLC" className={inputClass} />
            </Field>
            <Field label="Property Address">
              <input type="text" value={form.PropertyAddress ?? ''} onChange={(e) => handleChange('PropertyAddress', e.target.value)} placeholder="Street, City, State, ZIP" className={inputClass} />
            </Field>
            <Field label="Date Added to CAHP">
              <input type="date" value={form.DateAddedToCAHP ? new Date(form.DateAddedToCAHP).toISOString().slice(0, 10) : ''} onChange={(e) => handleChange('DateAddedToCAHP', e.target.value ? new Date(e.target.value).toISOString() : undefined)} className={`${inputClass} font-mono-data`} />
            </Field>
          </Section>
          <Section title="Location">
            <Field label="State" required>
              <select value={form.cahpState ?? ''} onChange={(e) => handleChange('cahpState', (e.target.value || undefined) as CahpState)} className={`${inputClass} bg-white font-mono-data`}>
                <option value="">— select —</option>
                {CHOICES.cahpState.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="County">
              <select value={form.cahpCounty ?? ''} onChange={(e) => handleChange('cahpCounty', e.target.value || undefined)} className={`${inputClass} bg-white`}>
                <option value="">— select —</option>
                {CHOICES.cahpCounty.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="Owner Group">
              <select value={form.cahpOwnerGroup ?? ''} onChange={(e) => handleChange('cahpOwnerGroup', (e.target.value || undefined) as OwnerGroup)} className={`${inputClass} bg-white`}>
                <option value="">— select —</option>
                {CHOICES.cahpOwnerGroup.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
          </Section>
        </div>
      )}

      {/* Step 2 — Program Details */}
      {step === 2 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Section title="Affordability Program">
            <Field label="AMI Program">
              <select value={form.AMIProgram ?? ''} onChange={(e) => handleChange('AMIProgram', (e.target.value || undefined) as AMIProgram)} className={`${inputClass} bg-white`}>
                <option value="">— select —</option>
                {CHOICES.AMIProgram.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="CAHP Language Added">
              <select value={form.CAHPLanguageAdded ?? ''} onChange={(e) => handleChange('CAHPLanguageAdded', (e.target.value || undefined) as CAHPLanguageStatus)} className={`${inputClass} bg-white`}>
                <option value="">— select —</option>
                {CHOICES.CAHPLanguageAdded.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="LURA Executed">
              <select value={form.LURAExecuted ?? ''} onChange={(e) => handleChange('LURAExecuted', (e.target.value || undefined) as LURAExecutedStatus)} className={`${inputClass} bg-white`}>
                <option value="">— select —</option>
                {CHOICES.LURAExecuted.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="OA Version">
              <input type="text" value={form.OpAgreementVersion ?? ''} onChange={(e) => handleChange('OpAgreementVersion', e.target.value)} placeholder="e.g., v2.1, Amended 2024" className={`${inputClass} font-mono-data`} />
            </Field>
          </Section>
          <Section title="Particulars">
            <Field label="Units">
              <input type="number" value={form.UnitCount ?? ''} onChange={(e) => handleChange('UnitCount', e.target.value === '' ? undefined : Number(e.target.value))} className={`${inputClass} font-mono-data`} min={0} />
            </Field>
            <Field label="DOR Account ID">
              <input type="text" value={form.DORAccountID ?? ''} onChange={(e) => handleChange('DORAccountID', e.target.value)} className={`${inputClass} font-mono-data`} />
            </Field>
            <Field label="Property Status">
              <select value={form.PropertyStatus ?? 'Pending'} onChange={(e) => handleChange('PropertyStatus', e.target.value as PropertyStatus)} className={`${inputClass} bg-white`}>
                {CHOICES.PropertyStatus.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="Verification">
              <select value={form.cahpVerificationStatus ?? 'Inherited - Unverified'} onChange={(e) => handleChange('cahpVerificationStatus', e.target.value as VerificationStatus)} className={`${inputClass} bg-white`}>
                {CHOICES.cahpVerificationStatus.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
          </Section>
          <Section title="Notes" fullWidth>
            <textarea value={form.PropertyNotes ?? ''} onChange={(e) => handleChange('PropertyNotes', e.target.value)} rows={3} className={`${inputClass} resize-y`} />
          </Section>
        </div>
      )}

      {/* Step 3 — Ownership Setup */}
      {step === 3 && (
        <div>
          <Section title="Ownership Members" fullWidth>
            {/* CAHP SC LLC pre-fill */}
            <div className={`p-3 rounded border ${cahpScLLC ? 'bg-teal-50 border-teal-200' : 'bg-yellow-50 border-yellow-200'}`}>
              <div className="flex items-start gap-2">
                <Icon name="alert" size={14} className={cahpScLLC ? 'text-teal-700' : 'text-yellow-700'} />
                <div className="text-sm">
                  {cahpScLLC ? (
                    <>
                      <strong>CAHP SC LLC</strong> will be auto-added as <strong>Managing Member at 0.01%</strong> on Finish (matched from Owners list).
                    </>
                  ) : (
                    <>
                      <strong>CAHP SC LLC</strong> not found in Owners list — won't be auto-added. <Link to="/owners/new" className="underline text-teal-700">Create it first</Link> if you want it included.
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* Additional members */}
            <div className="mt-4">
              <div className="flex items-center justify-between mb-2">
                <div className="text-sm font-semibold text-gray-700">Additional Members</div>
                <button
                  onClick={addOwnerEntry}
                  className="text-xs text-teal-700 hover:text-teal-900 font-medium flex items-center gap-1"
                >
                  <Icon name="plus" size={12} />
                  Add Member
                </button>
              </div>
              {ownerEntries.length === 0 ? (
                <p className="text-xs text-gray-500 italic">No additional members added. You can add them later from the property's Ownership tab.</p>
              ) : (
                <div className="space-y-2">
                  {ownerEntries.map((entry, idx) => (
                    <div key={idx} className="flex items-center gap-2 bg-gray-50 p-2 rounded">
                      <select
                        value={entry.ownerLookupId}
                        onChange={(e) => updateOwnerEntry(idx, { ownerLookupId: e.target.value })}
                        className={`${inputClass} bg-white flex-1`}
                      >
                        <option value="">— pick owner —</option>
                        {allOwners?.map((o) => (
                          <option key={o.id} value={o.id}>{o.fields.Title} {o.fields.OwnerType ? `(${o.fields.OwnerType})` : ''}</option>
                        ))}
                      </select>
                      <select
                        value={entry.role}
                        onChange={(e) => updateOwnerEntry(idx, { role: e.target.value as 'Managing Member' | 'Member' })}
                        className={`${inputClass} bg-white w-40`}
                      >
                        <option value="Member">Member</option>
                        <option value="Managing Member">Managing Member</option>
                      </select>
                      <input
                        type="number"
                        value={entry.percent}
                        onChange={(e) => updateOwnerEntry(idx, { percent: Number(e.target.value) })}
                        className={`${inputClass} w-24 font-mono-data text-right`}
                        min={0} max={100} step="0.01"
                        placeholder="%"
                      />
                      <button
                        onClick={() => removeOwnerEntry(idx)}
                        className="text-error hover:text-red-700 px-2"
                        title="Remove"
                      >×</button>
                    </div>
                  ))}
                </div>
              )}
              <div className="mt-3 text-xs text-gray-500">
                Total assigned: <span className="font-mono-data font-semibold">{totalOwnerPercent.toFixed(2)}%</span>
                {Math.abs(totalOwnerPercent - 100) > 0.01 && (
                  <span className="text-warning"> · should sum to 100% (or to class totals if multi-class OA)</span>
                )}
              </div>
            </div>
          </Section>
        </div>
      )}

      {/* Step 4 — Filing Config */}
      {step === 4 && (
        <Section title="CAHP Filing Configuration" fullWidth>
          <Field label="CAHP Fee Percent">
            <input
              type="number"
              value={filingConfig.cahpFeePercent}
              onChange={(e) => setFilingConfig({ ...filingConfig, cahpFeePercent: Number(e.target.value) })}
              className={`${inputClass} font-mono-data`}
              min={0} max={100} step="0.1"
            />
            <p className="text-xs text-gray-400 mt-1">Share of tax savings billed by CAHP. Typically 30–50%.</p>
          </Field>
          <Field label="First Filing Type">
            <select
              value={filingConfig.firstFilingType}
              onChange={(e) => setFilingConfig({ ...filingConfig, firstFilingType: e.target.value as 'Initial' | 'Annual' })}
              className={`${inputClass} bg-white`}
            >
              {CHOICES.FilingType.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <p className="text-xs text-gray-400 mt-1"><strong>Initial</strong> for the first CAHP filing on a property; <strong>Annual</strong> for ongoing.</p>
          </Field>
          <Field label="Tax Year">
            <select
              value={filingConfig.firstFilingTaxYear}
              onChange={(e) => setFilingConfig({ ...filingConfig, firstFilingTaxYear: e.target.value as CahpTaxYear })}
              className={`${inputClass} bg-white font-mono-data`}
            >
              {CHOICES.TaxYear.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="Annual Filing Required">
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={filingConfig.annualFilingRequired}
                onChange={(e) => setFilingConfig({ ...filingConfig, annualFilingRequired: e.target.checked })}
              />
              <span className="text-sm text-gray-700">Auto-create annual submittals each January 1</span>
            </label>
          </Field>
        </Section>
      )}

      {/* Step 5 — Initial Documents Preview */}
      {step === 5 && (
        <Section title="Initial Documents Outstanding Items" fullWidth>
          <p className="text-sm text-gray-600 mb-3">
            On Finish, the following <strong>{INITIAL_DOCS.length} Outstanding Items</strong> will be auto-created and linked to this property,
            all marked as <strong>Requested</strong>. You can edit, assign, or close them after creation from the Outstanding Items module.
          </p>
          <ul className="bg-gray-50 rounded-md p-3 space-y-1.5">
            {INITIAL_DOCS.map((doc, idx) => (
              <li key={idx} className="flex items-center gap-2 text-sm">
                <span className="inline-flex items-center justify-center w-5 h-5 rounded-full bg-teal-100 text-teal-700 text-xs font-bold flex-shrink-0">
                  {idx + 1}
                </span>
                <span className="text-gray-900">{doc.title}</span>
                <span className="text-xs text-gray-500 font-mono-data ml-auto">{doc.category}</span>
              </li>
            ))}
          </ul>
        </Section>
      )}

      {/* Step 6 — Review */}
      {step === 6 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card p-5">
          <h3 className="text-base font-semibold text-teal-700 mb-3">Review before creating</h3>

          <div className="space-y-4">
            <div>
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Property</div>
              <dl className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-1 text-sm">
                <ReviewRow label="Name" value={form.Title} required />
                <ReviewRow label="Legal Entity" value={form.LegalEntity} />
                <ReviewRow label="Address" value={form.PropertyAddress} />
                <ReviewRow label="State" value={form.cahpState} required mono />
                <ReviewRow label="County" value={form.cahpCounty} />
                <ReviewRow label="Owner Group" value={form.cahpOwnerGroup} />
                <ReviewRow label="AMI Program" value={form.AMIProgram} />
                <ReviewRow label="Units" value={form.UnitCount?.toString()} mono />
                <ReviewRow label="Status" value={form.PropertyStatus} />
                <ReviewRow label="DOR Account ID" value={form.DORAccountID} mono />
              </dl>
            </div>

            <div className="border-t border-gray-100 pt-4">
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Filing Config</div>
              <dl className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-1 text-sm">
                <ReviewRow label="CAHP Fee %" value={`${filingConfig.cahpFeePercent}%`} mono />
                <ReviewRow label="First Filing" value={`${filingConfig.firstFilingType} ${filingConfig.firstFilingTaxYear}`} />
                <ReviewRow label="Annual Required" value={filingConfig.annualFilingRequired ? 'Yes' : 'No'} />
              </dl>
            </div>

            <div className="border-t border-gray-100 pt-4">
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">
                Records that will be created
              </div>
              <ul className="text-sm space-y-1">
                <li>✓ <strong>1 Property</strong> record</li>
                <li>✓ <strong>1 Submittal</strong> ({filingConfig.firstFilingType} {filingConfig.firstFilingTaxYear}, status Draft)</li>
                <li>✓ <strong>{INITIAL_DOCS.length} Outstanding Items</strong> for initial document collection</li>
                <li>✓ <strong>{(cahpScLLC ? 1 : 0) + ownerEntries.length} Ownership row{((cahpScLLC ? 1 : 0) + ownerEntries.length) === 1 ? '' : 's'}</strong>
                  {cahpScLLC && ` (CAHP SC LLC at 0.01% Managing Member${ownerEntries.length > 0 ? ` + ${ownerEntries.length} additional` : ''})`}
                </li>
                <li className="pt-1 text-xs text-gray-500">All actions audit-logged</li>
              </ul>
            </div>
          </div>
        </div>
      )}

      {/* Nav */}
      <div className="mt-6 flex items-center justify-between">
        <button
          onClick={() => navigate('/properties')}
          disabled={saving}
          className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
        >
          Cancel
        </button>
        <div className="flex items-center gap-2">
          {step > 1 && (
            <button onClick={goBack} disabled={saving} className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50">
              ← Back
            </button>
          )}
          {step < 6 && (
            <button onClick={goNext} className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium">
              Next →
            </button>
          )}
          {step === 6 && (
            <button onClick={handleCreate} disabled={saving} className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50">
              {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
              {saving ? 'Creating…' : 'Finish & Create'}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, children }: { label: string; required?: boolean; children: React.ReactNode }) {
  return (
    <div className="flex items-start gap-3">
      <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1.5">
        {label}{required && <span className="text-error ml-0.5">*</span>}
      </label>
      <div className="flex-1">{children}</div>
    </div>
  );
}

function ReviewRow({ label, value, required, mono }: { label: string; value?: string; required?: boolean; mono?: boolean }) {
  const hasValue = value !== null && value !== undefined && value !== '';
  return (
    <div className="flex items-start gap-3 py-1">
      <dt className="text-sm text-gray-500 w-40 flex-shrink-0">{label}</dt>
      <dd className={`text-sm flex-1 ${mono ? 'font-mono-data' : ''} ${hasValue ? 'text-gray-900' : required ? 'text-error italic' : 'text-gray-300'}`}>
        {hasValue ? value : required ? 'missing' : '—'}
      </dd>
    </div>
  );
}

