import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  createListItem,
  LIST_NAMES,
  type PropertyFields,
  type AMIProgram,
  type CAHPLanguageStatus,
  type CahpState,
  type LURAExecutedStatus,
  type OwnerGroup,
  type PropertyStatus,
  type VerificationStatus,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { BreadcrumbBar, Section, SaveErrorBanner } from '../components/detail';

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

type WizardStep = 1 | 2 | 3;

const STEP_TITLES: Record<WizardStep, { title: string; subtitle: string }> = {
  1: { title: 'Identity & Location', subtitle: 'Who is this property and where is it?' },
  2: { title: 'Program Details', subtitle: 'How does it fit the CAHP affordability program?' },
  3: { title: 'Review & Create', subtitle: 'Double-check before writing to SharePoint.' },
};

export function PropertyNew() {
  const navigate = useNavigate();
  const [step, setStep] = useState<WizardStep>(1);
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
    if (err) {
      setValidationError(err);
      return;
    }
    if (step < 3) setStep((step + 1) as WizardStep);
  };

  const goBack = () => {
    setValidationError(null);
    if (step > 1) setStep((step - 1) as WizardStep);
  };

  const handleCreate = async () => {
    const err = validateStep(1);
    if (err) {
      setValidationError(err);
      setStep(1);
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // Drop empty/undefined values
      const fields: Record<string, unknown> = {};
      Object.entries(form).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') fields[k] = v;
      });
      const created = await createListItem<{ id: string }>(LIST_NAMES.Properties, fields);
      navigate(`/properties/${created.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  return (
    <div>
      <BreadcrumbBar parentLabel="Properties" parentTo="/properties" currentLabel="New Property" />

      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">New Property</h1>
        <p className="text-sm text-gray-500 mt-1">
          {STEP_TITLES[step].subtitle}
        </p>
      </div>

      {/* Stepper */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-4">
        <div className="flex items-center gap-2 sm:gap-4">
          {([1, 2, 3] as WizardStep[]).map((s, idx) => {
            const isCurrent = s === step;
            const isComplete = s < step;
            return (
              <div key={s} className="flex items-center gap-2 sm:gap-4 flex-1">
                <div className="flex items-center gap-2 min-w-0">
                  <div
                    className={`flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold ${
                      isCurrent
                        ? 'bg-teal-700 text-white'
                        : isComplete
                          ? 'bg-success text-white'
                          : 'bg-gray-100 text-gray-400'
                    }`}
                  >
                    {isComplete ? <Icon name="check" size={14} /> : s}
                  </div>
                  <div className="hidden sm:block min-w-0">
                    <div className={`text-xs font-semibold ${isCurrent ? 'text-teal-700' : isComplete ? 'text-success' : 'text-gray-400'}`}>
                      Step {s}
                    </div>
                    <div className={`text-sm truncate ${isCurrent ? 'text-gray-900 font-medium' : 'text-gray-500'}`}>
                      {STEP_TITLES[s].title}
                    </div>
                  </div>
                </div>
                {idx < 2 && <div className={`flex-1 h-0.5 ${isComplete ? 'bg-success' : 'bg-gray-200'}`} />}
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
            <FormField label="Property Name" required>
              <input
                type="text"
                value={form.Title ?? ''}
                onChange={(e) => handleChange('Title', e.target.value)}
                placeholder="e.g., 135 Oakwood Apartments"
                className={inputClass}
                autoFocus
              />
            </FormField>
            <FormField label="Legal Entity">
              <input
                type="text"
                value={form.LegalEntity ?? ''}
                onChange={(e) => handleChange('LegalEntity', e.target.value)}
                placeholder="e.g., 135 Oakwood LLC"
                className={inputClass}
              />
            </FormField>
            <FormField label="Property Address">
              <input
                type="text"
                value={form.PropertyAddress ?? ''}
                onChange={(e) => handleChange('PropertyAddress', e.target.value)}
                placeholder="Street address, city, state, ZIP"
                className={inputClass}
              />
            </FormField>
            <FormField label="Date Added to CAHP">
              <input
                type="date"
                value={form.DateAddedToCAHP ? new Date(form.DateAddedToCAHP).toISOString().slice(0, 10) : ''}
                onChange={(e) => handleChange('DateAddedToCAHP', e.target.value ? new Date(e.target.value).toISOString() : undefined)}
                className={`${inputClass} font-mono-data`}
              />
            </FormField>
          </Section>

          <Section title="Location">
            <FormField label="State" required>
              <select
                value={form.cahpState ?? ''}
                onChange={(e) => handleChange('cahpState', (e.target.value || undefined) as CahpState)}
                className={`${inputClass} bg-white font-mono-data`}
              >
                <option value="">— select —</option>
                {CHOICES.cahpState.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="County">
              <select
                value={form.cahpCounty ?? ''}
                onChange={(e) => handleChange('cahpCounty', e.target.value || undefined)}
                className={`${inputClass} bg-white`}
              >
                <option value="">— select —</option>
                {CHOICES.cahpCounty.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="Owner Group">
              <select
                value={form.cahpOwnerGroup ?? ''}
                onChange={(e) => handleChange('cahpOwnerGroup', (e.target.value || undefined) as OwnerGroup)}
                className={`${inputClass} bg-white`}
              >
                <option value="">— select —</option>
                {CHOICES.cahpOwnerGroup.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
          </Section>
        </div>
      )}

      {/* Step 2 — Program Details */}
      {step === 2 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <Section title="Affordability Program">
            <FormField label="AMI Program">
              <select
                value={form.AMIProgram ?? ''}
                onChange={(e) => handleChange('AMIProgram', (e.target.value || undefined) as AMIProgram)}
                className={`${inputClass} bg-white`}
              >
                <option value="">— select —</option>
                {CHOICES.AMIProgram.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="CAHP Language Added">
              <select
                value={form.CAHPLanguageAdded ?? ''}
                onChange={(e) => handleChange('CAHPLanguageAdded', (e.target.value || undefined) as CAHPLanguageStatus)}
                className={`${inputClass} bg-white`}
              >
                <option value="">— select —</option>
                {CHOICES.CAHPLanguageAdded.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="LURA Executed">
              <select
                value={form.LURAExecuted ?? ''}
                onChange={(e) => handleChange('LURAExecuted', (e.target.value || undefined) as LURAExecutedStatus)}
                className={`${inputClass} bg-white`}
              >
                <option value="">— select —</option>
                {CHOICES.LURAExecuted.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="Operating Agreement Version">
              <input
                type="text"
                value={form.OpAgreementVersion ?? ''}
                onChange={(e) => handleChange('OpAgreementVersion', e.target.value)}
                placeholder="e.g., v2.1, Amended 2024"
                className={`${inputClass} font-mono-data`}
              />
            </FormField>
          </Section>

          <Section title="Property Particulars">
            <FormField label="Units">
              <input
                type="number"
                value={form.UnitCount ?? ''}
                onChange={(e) => handleChange('UnitCount', e.target.value === '' ? undefined : Number(e.target.value))}
                placeholder="Total unit count"
                className={`${inputClass} font-mono-data`}
                min={0}
              />
            </FormField>
            <FormField label="DOR Account ID">
              <input
                type="text"
                value={form.DORAccountID ?? ''}
                onChange={(e) => handleChange('DORAccountID', e.target.value)}
                placeholder="SC DOR / NC DOR account number"
                className={`${inputClass} font-mono-data`}
              />
            </FormField>
            <FormField label="Property Status">
              <select
                value={form.PropertyStatus ?? 'Pending'}
                onChange={(e) => handleChange('PropertyStatus', e.target.value as PropertyStatus)}
                className={`${inputClass} bg-white`}
              >
                {CHOICES.PropertyStatus.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="Verification Status">
              <select
                value={form.cahpVerificationStatus ?? 'Inherited - Unverified'}
                onChange={(e) => handleChange('cahpVerificationStatus', e.target.value as VerificationStatus)}
                className={`${inputClass} bg-white`}
              >
                {CHOICES.cahpVerificationStatus.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </FormField>
          </Section>

          <Section title="Notes" fullWidth>
            <textarea
              value={form.PropertyNotes ?? ''}
              onChange={(e) => handleChange('PropertyNotes', e.target.value)}
              placeholder="Initial onboarding context — original owner, acquisition path, special considerations…"
              rows={4}
              className={`${inputClass} resize-y`}
            />
          </Section>
        </div>
      )}

      {/* Step 3 — Review */}
      {step === 3 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card p-5">
          <h3 className="text-base font-semibold text-teal-700 mb-4">Review before creating</h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-2 text-sm">
            <ReviewRow label="Property Name" value={form.Title} required />
            <ReviewRow label="Legal Entity" value={form.LegalEntity} />
            <ReviewRow label="Address" value={form.PropertyAddress} />
            <ReviewRow label="State" value={form.cahpState} required mono />
            <ReviewRow label="County" value={form.cahpCounty} />
            <ReviewRow label="Owner Group" value={form.cahpOwnerGroup} />
            <ReviewRow label="Date Added" value={form.DateAddedToCAHP ? new Date(form.DateAddedToCAHP).toLocaleDateString() : undefined} />
            <ReviewRow label="AMI Program" value={form.AMIProgram} />
            <ReviewRow label="CAHP Language Added" value={form.CAHPLanguageAdded} />
            <ReviewRow label="LURA Executed" value={form.LURAExecuted} />
            <ReviewRow label="OA Version" value={form.OpAgreementVersion} mono />
            <ReviewRow label="Units" value={form.UnitCount?.toString()} mono />
            <ReviewRow label="DOR Account ID" value={form.DORAccountID} mono />
            <ReviewRow label="Property Status" value={form.PropertyStatus} />
            <ReviewRow label="Verification Status" value={form.cahpVerificationStatus} />
          </div>
          {form.PropertyNotes && (
            <div className="mt-4 pt-4 border-t border-gray-100">
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Notes</div>
              <p className="text-sm text-gray-700 whitespace-pre-wrap">{form.PropertyNotes}</p>
            </div>
          )}

          <div className="mt-6 bg-blue-50 border border-blue-200 rounded-lg p-3 flex items-start gap-2">
            <Icon name="alert" size={14} className="text-blue-700 flex-shrink-0 mt-0.5" />
            <p className="text-xs text-blue-900">
              Clicking <strong>Create Property</strong> writes a new record to Properties Registry in SharePoint
              and logs the action to the Audit Log. You'll land on the new property's detail page.
            </p>
          </div>
        </div>
      )}

      {/* Navigation */}
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
            <button
              onClick={goBack}
              disabled={saving}
              className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
            >
              ← Back
            </button>
          )}
          {step < 3 && (
            <button
              onClick={goNext}
              className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium transition-colors"
            >
              Next →
            </button>
          )}
          {step === 3 && (
            <button
              onClick={handleCreate}
              disabled={saving}
              className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50"
            >
              {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
              {saving ? 'Creating…' : 'Create Property'}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function FormField({ label, required, children }: { label: string; required?: boolean; children: React.ReactNode }) {
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
        {hasValue ? value : required ? 'missing — required' : '—'}
      </dd>
    </div>
  );
}
