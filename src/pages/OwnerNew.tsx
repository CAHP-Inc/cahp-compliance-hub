import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  createListItem,
  LIST_NAMES,
  type OwnerFields,
  type OwnerType,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { BreadcrumbBar, SaveErrorBanner } from '../components/detail';

type WizardStep = 1 | 2 | 3;

const STEP_INFO: Record<WizardStep, { title: string; subtitle: string }> = {
  1: { title: 'Owner Type', subtitle: 'What kind of owner is this?' },
  2: { title: 'Entity Info', subtitle: 'Legal identity and contact info.' },
  3: { title: 'Review & Create', subtitle: 'Confirm before writing to SharePoint.' },
};

const TYPE_CARDS: { value: OwnerType; title: string; description: string; icon: string }[] = [
  {
    value: 'Individual',
    title: 'Individual',
    description: 'A natural person — Maksim Grushkovskiy, Brandy Turner. For joint owners (married couples), create one Individual record per person.',
    icon: '👤',
  },
  {
    value: 'LLC',
    title: 'LLC',
    description: 'A limited liability company — VanRock Holdings LLC, 135 Oakwood LLC. Can have members (other LLCs or Individuals).',
    icon: '🏢',
  },
  {
    value: 'Nonprofit',
    title: 'Nonprofit',
    description: 'A 501(c)(3) or similar nonprofit entity — CAHP SC LLC (despite the LLC name, treated as nonprofit for CAHP).',
    icon: '🏛️',
  },
];

export function OwnerNew() {
  const navigate = useNavigate();
  const [step, setStep] = useState<WizardStep>(1);
  const [form, setForm] = useState<Partial<OwnerFields>>({
    Title: '',
    OwnerType: undefined,
    OwnerState: '',
    TaxID: '',
    ContactEmail: '',
    OwnerNotes: '',
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);

  const handleChange = <K extends keyof OwnerFields>(field: K, value: OwnerFields[K]) => {
    setForm((prev) => ({ ...prev, [field]: value }));
    setValidationError(null);
  };

  const validateStep = (s: WizardStep): string | null => {
    if (s === 1 && !form.OwnerType) return 'Choose an owner type to continue.';
    if (s === 2) {
      if (!form.Title || form.Title.trim() === '') return 'Legal name is required.';
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
    const e1 = validateStep(1);
    const e2 = validateStep(2);
    if (e1) { setValidationError(e1); setStep(1); return; }
    if (e2) { setValidationError(e2); setStep(2); return; }

    setSaving(true);
    setError(null);
    try {
      const fields: Record<string, unknown> = {};
      Object.entries(form).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') fields[k] = v;
      });
      const created = await createListItem<{ id: string }>(LIST_NAMES.Owners, fields);
      navigate(`/owners/${created.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  return (
    <div>
      <BreadcrumbBar parentLabel="Owners" parentTo="/owners" currentLabel="New Owner" />

      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">New Owner</h1>
        <p className="text-sm text-gray-500 mt-1">{STEP_INFO[step].subtitle}</p>
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
                      {STEP_INFO[s].title}
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

      {/* Step 1 — Type selection (card picker) */}
      {step === 1 && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {TYPE_CARDS.map((card) => (
            <button
              key={card.value}
              onClick={() => handleChange('OwnerType', card.value)}
              className={`bg-white border-2 rounded-lg p-5 text-left transition-all hover:shadow-card-hover ${
                form.OwnerType === card.value
                  ? 'border-teal-700 shadow-card'
                  : 'border-gray-200 hover:border-gray-300'
              }`}
            >
              <div className="text-3xl mb-3">{card.icon}</div>
              <h3 className="text-base font-bold text-teal-900 mb-2">{card.title}</h3>
              <p className="text-xs text-gray-600 leading-relaxed">{card.description}</p>
            </button>
          ))}
        </div>
      )}

      {/* Step 2 — Entity Info */}
      {step === 2 && (
        <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
          <div className="space-y-3">
            <Field label="Legal Name" required>
              <input
                type="text"
                value={form.Title ?? ''}
                onChange={(e) => handleChange('Title', e.target.value)}
                placeholder={
                  form.OwnerType === 'Individual'
                    ? 'e.g., Maksim Grushkovskiy'
                    : form.OwnerType === 'LLC'
                      ? 'e.g., VanRock Holdings LLC'
                      : 'e.g., CAHP SC LLC'
                }
                className={inputClass}
                autoFocus
              />
            </Field>
            <Field label={form.OwnerType === 'Individual' ? 'State of Residence' : 'State of Formation'}>
              <input
                type="text"
                value={form.OwnerState ?? ''}
                onChange={(e) => handleChange('OwnerState', e.target.value)}
                placeholder="e.g., SC, NC, NY"
                className={`${inputClass} font-mono-data`}
              />
            </Field>
            <Field label={form.OwnerType === 'Individual' ? 'SSN (optional)' : 'EIN (optional)'}>
              <input
                type="text"
                value={form.TaxID ?? ''}
                onChange={(e) => handleChange('TaxID', e.target.value)}
                placeholder={form.OwnerType === 'Individual' ? 'XXX-XX-XXXX' : 'XX-XXXXXXX'}
                className={`${inputClass} font-mono-data`}
              />
              <p className="text-xs text-gray-400 mt-1">
                Stored as-is in SharePoint; masked in app display. For real encryption, use a password vault and store only an identifier here.
              </p>
            </Field>
            <Field label="Contact Email">
              <input
                type="email"
                value={form.ContactEmail ?? ''}
                onChange={(e) => handleChange('ContactEmail', e.target.value)}
                placeholder="primary@example.com"
                className={inputClass}
              />
            </Field>
            <Field label="Notes">
              <textarea
                value={form.OwnerNotes ?? ''}
                onChange={(e) => handleChange('OwnerNotes', e.target.value)}
                rows={3}
                placeholder="Background, special considerations, communication preferences…"
                className={`${inputClass} resize-y`}
              />
            </Field>
          </div>
        </div>
      )}

      {/* Step 3 — Review */}
      {step === 3 && (
        <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
          <h3 className="text-base font-semibold text-teal-700 mb-4">Review before creating</h3>
          <dl className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-2 text-sm">
            <Row label="Owner Type" value={form.OwnerType} required />
            <Row label="Legal Name" value={form.Title} required />
            <Row label="State" value={form.OwnerState} mono />
            <Row label="Tax ID" value={form.TaxID ? '••••••' + form.TaxID.slice(-4) : undefined} mono />
            <Row label="Contact Email" value={form.ContactEmail} />
          </dl>
          {form.OwnerNotes && (
            <div className="mt-4 pt-4 border-t border-gray-100">
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Notes</div>
              <p className="text-sm text-gray-700 whitespace-pre-wrap">{form.OwnerNotes}</p>
            </div>
          )}
          <div className="mt-6 bg-blue-50 border border-blue-200 rounded-lg p-3 flex items-start gap-2">
            <Icon name="alert" size={14} className="text-blue-700 flex-shrink-0 mt-0.5" />
            <p className="text-xs text-blue-900">
              Clicking <strong>Create Owner</strong> writes a new record to the Owners list and lands
              you on the new owner's detail page, where you can add Members (if LLC/Nonprofit) or
              link them to properties.
            </p>
          </div>
        </div>
      )}

      {/* Nav */}
      <div className="mt-6 flex items-center justify-between">
        <button
          onClick={() => navigate('/owners')}
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
              {saving ? 'Creating…' : 'Create Owner'}
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

function Row({ label, value, required, mono }: { label: string; value?: string; required?: boolean; mono?: boolean }) {
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
