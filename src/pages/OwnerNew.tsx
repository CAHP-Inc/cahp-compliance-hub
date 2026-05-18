import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  createListItem,
  useSharePointList,
  LIST_NAMES,
  type Owner,
  type OwnerFields,
  type OwnerType,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { BreadcrumbBar, SaveErrorBanner } from '../components/detail';

type WizardStep = 1 | 2 | 3 | 4;

const STEP_INFO: Record<WizardStep, { title: string; subtitle: string }> = {
  1: { title: 'Type Selection', subtitle: 'What kind of owner is this?' },
  2: { title: 'Entity Info', subtitle: 'Legal identity and contact info.' },
  3: { title: 'Members (optional)', subtitle: 'Add members of this entity now — or later from its detail page.' },
  4: { title: 'Review & Create', subtitle: 'Confirm before writing to SharePoint.' },
};

const TYPE_CARDS: { value: OwnerType; title: string; description: string; icon: string }[] = [
  {
    value: 'Individual',
    title: 'Individual',
    description: 'A natural person. For joint owners (married couples), create one Individual record per person — keeps disbursement tracking clean.',
    icon: '👤',
  },
  {
    value: 'LLC',
    title: 'LLC',
    description: 'A limited liability company — including single-member LLCs wholly owned by a nonprofit (e.g., CAHP SC, LLC). Use LLC even if the LLC is tax-exempt by virtue of its nonprofit owner.',
    icon: '🏢',
  },
  {
    value: 'Nonprofit',
    title: 'Nonprofit Corporation',
    description: 'A 501(c)(3) or similar nonprofit CORPORATION itself (e.g., Carolina Affordable Housing Project, Inc.). Subsidiary LLCs owned by the nonprofit are still LLCs — pick LLC for those.',
    icon: '🏛️',
  },
  {
    value: 'Trust',
    title: 'Trust',
    description: 'A trust holding interests on behalf of beneficiaries (e.g., a family trust with trustees).',
    icon: '🛡️',
  },
  {
    value: 'Corporation',
    title: 'Corporation',
    description: 'A standard for-profit corporation (Inc.) — distinct from an LLC.',
    icon: '🏛',
  },
];

interface MemberDraft {
  ownerLookupId: string;
  role: 'Managing Member' | 'Member';
  percent: number;
}

export function OwnerNew() {
  const navigate = useNavigate();

  // Branching: Individual skips step 3 (Members); LLC/Nonprofit goes 1 → 2 → 3 → 4
  const [step, setStep] = useState<WizardStep>(1);

  const { data: allOwners } = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const [form, setForm] = useState<Partial<OwnerFields>>({
    Title: '',
    OwnerType: undefined,
    OwnerState: '',
    TaxID: '',
    ContactEmail: '',
    OwnerNotes: '',
  });

  const [members, setMembers] = useState<MemberDraft[]>([]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [validationError, setValidationError] = useState<string | null>(null);

  const hasMembersStep = form.OwnerType === 'LLC' || form.OwnerType === 'Nonprofit';

  // Visible steps depending on owner type
  const visibleSteps: WizardStep[] = useMemo(() => {
    return hasMembersStep ? [1, 2, 3, 4] : [1, 2, 4];
  }, [hasMembersStep]);

  const handleChange = <K extends keyof OwnerFields>(field: K, value: OwnerFields[K]) => {
    setForm((prev) => ({ ...prev, [field]: value }));
    setValidationError(null);
  };

  const validateStep = (s: WizardStep): string | null => {
    if (s === 1 && !form.OwnerType) return 'Choose an owner type to continue.';
    if (s === 2 && (!form.Title || form.Title.trim() === '')) return 'Legal name is required.';
    if (s === 3) {
      // Members are optional — but if any are added, validate each
      for (const m of members) {
        if (!m.ownerLookupId) return 'Every added member must have an owner selected.';
      }
    }
    return null;
  };

  const goNext = () => {
    const err = validateStep(step);
    if (err) { setValidationError(err); return; }
    // Find next visible step
    const currentIdx = visibleSteps.indexOf(step);
    if (currentIdx >= 0 && currentIdx < visibleSteps.length - 1) {
      setStep(visibleSteps[currentIdx + 1]);
    }
  };

  const goBack = () => {
    setValidationError(null);
    const currentIdx = visibleSteps.indexOf(step);
    if (currentIdx > 0) setStep(visibleSteps[currentIdx - 1]);
  };

  const handleCreate = async () => {
    // Re-validate all visible steps
    for (const s of visibleSteps) {
      const err = validateStep(s);
      if (err) { setValidationError(err); setStep(s); return; }
    }

    setSaving(true);
    setError(null);
    try {
      // 1. Create the Owner record
      const fields: Record<string, unknown> = {};
      Object.entries(form).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v !== '') fields[k] = v;
      });
      const created = await createListItem<{ id: string }>(LIST_NAMES.Owners, fields);

      // 2. Create member ownership rows (if any) — these are entity-to-entity
      // The new owner is the PARENT; each member has OwnerLookupId = their owner, ParentOwnerLookupId = the new owner
      for (const m of members) {
        const member = allOwners?.find((o) => String(o.id) === String(m.ownerLookupId));
        if (!member) continue;
        try {
          await createListItem(LIST_NAMES.Ownership, {
            Title: member.fields.Title,
            OwnerLookupId: m.ownerLookupId,
            ParentOwnerLookupId: created.id,
            RelationshipType: m.role,
            OwnershipPercent: m.percent,
            SourceDocument: 'Owner Creation Wizard',
          });
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn(`Member row for ${member.fields.Title} failed:`, e);
        }
      }

      navigate(`/owners/${created.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  const addMember = () => {
    setMembers((prev) => [...prev, { ownerLookupId: '', role: 'Member', percent: 0 }]);
  };

  const updateMember = (idx: number, patch: Partial<MemberDraft>) => {
    setMembers((prev) => prev.map((m, i) => (i === idx ? { ...m, ...patch } : m)));
  };

  const removeMember = (idx: number) => {
    setMembers((prev) => prev.filter((_, i) => i !== idx));
  };

  const totalPercent = members.reduce((sum, m) => sum + (m.percent || 0), 0);

  // Pretty-print visible step number (1-indexed within the visible flow)
  const stepDisplayNumber = visibleSteps.indexOf(step) + 1;
  const totalVisibleSteps = visibleSteps.length;
  const isLastStep = step === visibleSteps[visibleSteps.length - 1];

  return (
    <div>
      <BreadcrumbBar parentLabel="Owners" parentTo="/owners" currentLabel="New Owner" />

      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">New Owner</h1>
        <p className="text-sm text-gray-500 mt-1">{STEP_INFO[step].subtitle}</p>
      </div>

      {/* Stepper — shows only visible steps */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-4">
        <div className="flex items-center gap-2 sm:gap-4">
          {visibleSteps.map((s, idx) => {
            const isCurrent = s === step;
            const sIdx = visibleSteps.indexOf(s);
            const stepIdx = visibleSteps.indexOf(step);
            const isComplete = sIdx < stepIdx;
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
                    {isComplete ? <Icon name="check" size={14} /> : idx + 1}
                  </div>
                  <div className="hidden sm:block min-w-0">
                    <div className={`text-xs font-semibold ${isCurrent ? 'text-teal-700' : isComplete ? 'text-success' : 'text-gray-400'}`}>
                      Step {idx + 1}
                    </div>
                    <div className={`text-sm truncate ${isCurrent ? 'text-gray-900 font-medium' : 'text-gray-500'}`}>
                      {STEP_INFO[s].title}
                    </div>
                  </div>
                </div>
                {idx < visibleSteps.length - 1 && (
                  <div className={`flex-1 h-0.5 ${isComplete ? 'bg-success' : 'bg-gray-200'}`} />
                )}
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

      {/* Step 1 — Type Selection */}
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
                    ? 'e.g., Stan Gendlin'
                    : form.OwnerType === 'LLC'
                      ? 'e.g., VanRock Holdings LLC'
                      : 'e.g., Carolina Affordable Housing Project'
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

      {/* Step 3 — Members (LLC/Nonprofit only) */}
      {step === 3 && hasMembersStep && (
        <div className="bg-white border border-gray-200 rounded-lg p-5 shadow-card">
          <p className="text-sm text-gray-600 mb-4">
            Add members of <strong>{form.Title || 'this entity'}</strong> now, or skip and add them later from the owner detail page.
            Each member must be an existing Owner record. <button
              type="button"
              onClick={() => navigate('/owners/new')}
              className="text-teal-700 hover:text-teal-900 font-medium underline"
            >Create one first</button> if it doesn't exist yet.
          </p>

          <div className="flex items-center justify-between mb-2">
            <div className="text-sm font-semibold text-gray-700">Members</div>
            <button
              onClick={addMember}
              className="text-xs text-teal-700 hover:text-teal-900 font-medium flex items-center gap-1"
            >
              <Icon name="plus" size={12} />
              Add Member
            </button>
          </div>

          {members.length === 0 ? (
            <div className="bg-gray-50 rounded-md p-4 text-center">
              <p className="text-sm text-gray-500 italic">
                No members added. That's fine — you can add them later from <strong>{form.Title || 'this entity'}</strong>'s detail page.
              </p>
            </div>
          ) : (
            <div className="space-y-2">
              {members.map((m, idx) => (
                <div key={idx} className="flex items-center gap-2 bg-gray-50 p-2 rounded">
                  <select
                    value={m.ownerLookupId}
                    onChange={(e) => updateMember(idx, { ownerLookupId: e.target.value })}
                    className={`${inputClass} bg-white flex-1`}
                  >
                    <option value="">— pick owner —</option>
                    {allOwners?.map((o) => (
                      <option key={o.id} value={o.id}>
                        {o.fields.Title} {o.fields.OwnerType ? `(${o.fields.OwnerType})` : ''}
                      </option>
                    ))}
                  </select>
                  <select
                    value={m.role}
                    onChange={(e) => updateMember(idx, { role: e.target.value as 'Managing Member' | 'Member' })}
                    className={`${inputClass} bg-white w-40`}
                  >
                    <option value="Member">Member</option>
                    <option value="Managing Member">Managing Member</option>
                  </select>
                  <input
                    type="number"
                    value={m.percent}
                    onChange={(e) => updateMember(idx, { percent: Number(e.target.value) })}
                    className={`${inputClass} w-24 font-mono-data text-right`}
                    min={0} max={100} step="0.01"
                    placeholder="%"
                  />
                  <button
                    onClick={() => removeMember(idx)}
                    className="text-error hover:text-red-700 px-2"
                    title="Remove"
                  >×</button>
                </div>
              ))}
            </div>
          )}

          {members.length > 0 && (
            <div className="mt-3 text-xs text-gray-500">
              Total: <span className="font-mono-data font-semibold">{totalPercent.toFixed(2)}%</span>
              {Math.abs(totalPercent - 100) > 0.01 && (
                <span className="text-warning"> · members should sum to 100%</span>
              )}
            </div>
          )}
        </div>
      )}

      {/* Step 4 — Review */}
      {step === 4 && (
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
          {hasMembersStep && members.length > 0 && (
            <div className="mt-4 pt-4 border-t border-gray-100">
              <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">
                {members.length} Member{members.length === 1 ? '' : 's'} will be added
              </div>
              <ul className="space-y-1 text-sm">
                {members.map((m, idx) => {
                  const owner = allOwners?.find((o) => String(o.id) === String(m.ownerLookupId));
                  return (
                    <li key={idx} className="flex items-center gap-2">
                      <span className="font-medium">{owner?.fields.Title ?? '(missing)'}</span>
                      <span className="text-xs text-gray-500">— {m.role}, {m.percent}%</span>
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
          <div className="mt-6 bg-blue-50 border border-blue-200 rounded-lg p-3 flex items-start gap-2">
            <Icon name="alert" size={14} className="text-blue-700 flex-shrink-0 mt-0.5" />
            <p className="text-xs text-blue-900">
              Clicking <strong>Create Owner</strong> writes the new record and{' '}
              {hasMembersStep && members.length > 0 ? `${members.length} ownership row${members.length === 1 ? '' : 's'} for members` : 'lands you on the owner detail page'}.
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
          <span className="text-xs text-gray-400 hidden sm:inline">Step {stepDisplayNumber} of {totalVisibleSteps}</span>
          {stepDisplayNumber > 1 && (
            <button
              onClick={goBack}
              disabled={saving}
              className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
            >
              ← Back
            </button>
          )}
          {!isLastStep && (
            <button
              onClick={goNext}
              className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium transition-colors"
            >
              Next →
            </button>
          )}
          {isLastStep && (
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
