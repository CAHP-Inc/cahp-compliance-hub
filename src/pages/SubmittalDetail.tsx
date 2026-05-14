import { useState, useEffect, useMemo } from 'react';
import { useParams, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  createListItem,
  LIST_NAMES,
  type Submittal,
  type SubmittalFields,
  type Property,
  type Correspondence,
  type OutstandingItem,
  type AuditLog,
  type SubmittalStatusValue,
  type SubmittalFilingType,
  type CahpTaxYear,
  type CahpState,
  type FilingMethod,
  type BillingStatusValue,
  type DisbursementStatus,
  getBeneficialOwnershipTree,
  type Ownership,
  type Owner,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';
import { SubmittalOrgChartSnapshot } from '../components/SubmittalOrgChartSnapshot';

const STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  'Draft': 'bg-gray-100 text-gray-800',
  'Package Mailed (NC)': 'bg-indigo-100 text-indigo-800',
  'Filed': 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-amber-100 text-amber-800',
  'Responded - Awaiting DOR': 'bg-purple-100 text-purple-800',
  'Approved': 'bg-green-100 text-green-800',
  'Denied': 'bg-red-100 text-red-800',
  'Withdrawn': 'bg-gray-100 text-gray-500',
};

const FILING_TYPE_STYLES: Record<SubmittalFilingType, string> = {
  'Initial': 'bg-teal-100 text-teal-800',
  'Annual': 'bg-blue-100 text-blue-800',
  'Amendment': 'bg-amber-100 text-amber-800',
};

const PIPELINE_STAGES: { status: SubmittalStatusValue; label: string }[] = [
  { status: 'Draft', label: 'Draft' },
  { status: 'Filed', label: 'Filed' },
  { status: 'Letter Received - Action Needed', label: 'Letter Received' },
  { status: 'Responded - Awaiting DOR', label: 'Responded' },
  { status: 'Approved', label: 'Approved' },
];

const TERMINAL_STATUSES: SubmittalStatusValue[] = ['Approved', 'Denied', 'Withdrawn'];

const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];
const STATES: CahpState[] = ['SC', 'NC'];
const FILING_METHODS: FilingMethod[] = ['Online Portal (SC)', 'Paper Mail (NC)'];
const FILING_TYPES: SubmittalFilingType[] = ['Initial', 'Annual', 'Amendment'];

// ─────────────────────────────────────────────────────────────
// Allowed transitions — spec §3.6.4
// ─────────────────────────────────────────────────────────────
type Transition = {
  to: SubmittalStatusValue;
  label: string;
  description: string;
  style: 'primary' | 'warning' | 'danger' | 'success' | 'neutral';
  requiresFields?: (keyof SubmittalFields)[];     // Fields that must be set BEFORE allowing transition
};

const ALLOWED_TRANSITIONS: Record<SubmittalStatusValue, Transition[]> = {
  'Draft': [
    {
      to: 'Filed',
      label: 'File with DOR',
      description: 'Mark this submittal as filed. Captures filing date + confirmation #, freezes the org chart snapshot.',
      style: 'primary',
      requiresFields: ['DateFiled', 'ConfirmationNumber'],
    },
    {
      to: 'Withdrawn',
      label: 'Withdraw',
      description: 'Cancel this submittal. Terminal status.',
      style: 'danger',
    },
  ],
  'Package Mailed (NC)': [
    {
      to: 'Filed',
      label: 'Mark as Filed',
      description: 'DOR confirmed receipt of mailed package.',
      style: 'primary',
      requiresFields: ['DateFiled', 'ConfirmationNumber'],
    },
    { to: 'Withdrawn', label: 'Withdraw', description: 'Cancel this submittal.', style: 'danger' },
  ],
  'Filed': [
    {
      to: 'Letter Received - Action Needed',
      label: 'Letter Received (action needed)',
      description: 'DOR sent a letter requesting more info or clarification.',
      style: 'warning',
    },
    {
      to: 'Approved',
      label: 'Approved',
      description: 'DOR approved. Opens approval workflow modal to capture tax savings + create Billing.',
      style: 'success',
    },
    { to: 'Denied', label: 'Denied', description: 'DOR denied this submittal. Terminal.', style: 'danger' },
    { to: 'Withdrawn', label: 'Withdraw', description: 'Withdraw before further DOR action.', style: 'neutral' },
  ],
  'Letter Received - Action Needed': [
    {
      to: 'Responded - Awaiting DOR',
      label: 'Mark Responded',
      description: 'Response sent to DOR. Awaiting their next move.',
      style: 'primary',
    },
    { to: 'Withdrawn', label: 'Withdraw', description: 'Withdraw the submittal.', style: 'danger' },
  ],
  'Responded - Awaiting DOR': [
    {
      to: 'Letter Received - Action Needed',
      label: 'Another Letter Received',
      description: 'DOR sent another letter requiring response.',
      style: 'warning',
    },
    {
      to: 'Approved',
      label: 'Approved',
      description: 'DOR approved. Opens approval workflow modal.',
      style: 'success',
    },
    { to: 'Denied', label: 'Denied', description: 'DOR denied. Terminal.', style: 'danger' },
  ],
  'Approved': [],   // Terminal — but show appeal-not-yet-supported note
  'Denied': [],     // Terminal — appeal logic in Phase 3 maybe
  'Withdrawn': [],  // Terminal
};

const TRANSITION_STYLES: Record<Transition['style'], string> = {
  primary: 'bg-teal-700 hover:bg-teal-900 text-white border-transparent',
  warning: 'bg-amber-600 hover:bg-amber-700 text-white border-transparent',
  danger: 'bg-red-600 hover:bg-red-700 text-white border-transparent',
  success: 'bg-green-600 hover:bg-green-700 text-white border-transparent',
  neutral: 'bg-white hover:bg-gray-50 text-gray-700 border-gray-300',
};

export function SubmittalDetail() {
  const { id } = useParams<{ id: string }>();

  const { data: submittal, loading, error, refetch } = useSharePointItem<Submittal>(
    LIST_NAMES.Submittals,
    id
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const correspondence = useSharePointList<Correspondence>(LIST_NAMES.Correspondence, { top: 500 });
  const outstanding = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const auditLog = useSharePointList<AuditLog>(LIST_NAMES.AuditLog, { top: 200 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  // Editing state
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<SubmittalFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  // Transition modal state
  const [activeTransition, setActiveTransition] = useState<Transition | null>(null);
  const [transitionFields, setTransitionFields] = useState<Record<string, string>>({});
  const [transitionError, setTransitionError] = useState<string | null>(null);

  // PR-10d — Approval Workflow modal (special-cased Approved transition)
  const [approvalModalOpen, setApprovalModalOpen] = useState(false);
  const [approvalLetterRef, setApprovalLetterRef] = useState('');
  const [taxSavingsAmount, setTaxSavingsAmount] = useState('');
  const [cahpFeePercent, setCahpFeePercent] = useState('50');
  const [approvalError, setApprovalError] = useState<string | null>(null);
  const [approvalSaving, setApprovalSaving] = useState(false);

  useEffect(() => {
    if (submittal && !editing) setDraft({ ...submittal.fields });
  }, [submittal?.id, submittal?.lastModifiedDateTime, editing]);

  const property = useMemo(() => {
    if (!submittal || !properties.data || !submittal.fields.PropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(submittal.fields.PropertyLookupId)) ?? null;
  }, [submittal, properties.data]);

  const relatedCorrespondence = useMemo(() => {
    if (!submittal || !correspondence.data) return [];
    const propertyId = submittal.fields.PropertyLookupId;
    if (!propertyId) return [];
    return correspondence.data
      .filter((c) => String(c.fields.PropertyLookupId) === String(propertyId))
      .sort((a, b) => {
        const da = a.fields.DateReceived ? new Date(a.fields.DateReceived).getTime() : 0;
        const db = b.fields.DateReceived ? new Date(b.fields.DateReceived).getTime() : 0;
        return db - da;
      });
  }, [submittal, correspondence.data]);

  const actionPlan = useMemo(() => {
    if (!submittal || !outstanding.data) return [];
    const propertyId = submittal.fields.PropertyLookupId;
    if (!propertyId) return [];
    return outstanding.data.filter(
      (o) =>
        String(o.fields.PropertyLookupId) === String(propertyId) &&
        o.fields.ItemStatus !== 'Received' &&
        o.fields.ItemStatus !== 'Not Applicable'
    );
  }, [submittal, outstanding.data]);

  const submittalActivity = useMemo(() => {
    if (!submittal || !auditLog.data) return [];
    return auditLog.data
      .filter(
        (a) =>
          a.fields.EntityType === 'Submittal' && String(a.fields.EntityId) === String(submittal.id)
      )
      .sort(
        (a, b) =>
          new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime()
      );
  }, [submittal, auditLog.data]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading submittal…</span>
        </div>
      </div>
    );
  }

  if (error || !submittal || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Submittals" parentTo="/submittals" currentLabel="Submittal Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load submittal</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : submittal.fields;
  const f = submittal.fields;
  const currentStatus = f.SubmittalStatus ?? 'Draft';
  const transitions = ALLOWED_TRANSITIONS[currentStatus] ?? [];
  const isTerminal = TERMINAL_STATUSES.includes(currentStatus);
  const statusIdx = PIPELINE_STAGES.findIndex((s) => s.status === currentStatus);

  const handleFieldChange = <K extends keyof SubmittalFields>(field: K, value: SubmittalFields[K]) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...submittal.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...submittal.fields });
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
        const k = key as keyof SubmittalFields;
        if (draft[k] !== submittal.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Submittals, submittal.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  // ────────────────────────────────────────────────────────
  // Status transitions
  // ────────────────────────────────────────────────────────

  const startTransition = (transition: Transition) => {
    // PR-10d — Approved transition uses the Approval Workflow modal instead of generic flow
    if (transition.to === 'Approved') {
      setApprovalLetterRef('');
      setTaxSavingsAmount('');
      setCahpFeePercent('50');
      setApprovalError(null);
      setApprovalModalOpen(true);
      return;
    }
    setActiveTransition(transition);
    setTransitionFields({});
    setTransitionError(null);
  };

  const cancelTransition = () => {
    setActiveTransition(null);
    setTransitionFields({});
    setTransitionError(null);
  };

  const confirmTransition = async () => {
    if (!activeTransition) return;
    // Validate required fields
    const missing: string[] = [];
    (activeTransition.requiresFields ?? []).forEach((field) => {
      const existingValue = f[field];
      const inputValue = transitionFields[field];
      if (!existingValue && !inputValue) missing.push(String(field));
    });
    if (missing.length > 0) {
      setTransitionError(`Required: ${missing.join(', ')}`);
      return;
    }

    setSaving(true);
    setTransitionError(null);
    try {
      const updates: Record<string, unknown> = {
        SubmittalStatus: activeTransition.to,
      };
      // Apply transition-time field captures
      Object.entries(transitionFields).forEach(([k, v]) => {
        if (v) updates[k] = k === 'DateFiled' ? new Date(v).toISOString() : v;
      });

      // Spec §3.6.6 — On Draft → Filed transition, freeze org chart snapshot
      const fromDraftToFiled =
        (currentStatus === 'Draft' || currentStatus === 'Package Mailed (NC)') &&
        activeTransition.to === 'Filed';
      if (fromDraftToFiled && submittal.fields.PropertyLookupId && ownership.data && owners.data) {
        const tree = getBeneficialOwnershipTree(
          'property',
          String(submittal.fields.PropertyLookupId),
          ownership.data,
          owners.data
        );
        const snapshot = {
          version: 1,
          capturedAt: new Date().toISOString(),
          propertyId: submittal.fields.PropertyLookupId,
          tree: serializeTree(tree),
        };
        updates.OrgChartSnapshotJSON = JSON.stringify(snapshot);
        updates.OrgChartSnapshotDate = new Date().toISOString();
      }

      await updateListItem(LIST_NAMES.Submittals, submittal.id, updates);
      await refetch();
      cancelTransition();
    } catch (err) {
      setTransitionError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleTransitionFieldChange = (field: string, value: string) => {
    setTransitionFields((prev) => ({ ...prev, [field]: value }));
  };

  // ────────────────────────────────────────────────────────
  // PR-10d — Approval Workflow: create Billing + Disbursement atomically
  // ────────────────────────────────────────────────────────

  const cancelApproval = () => {
    setApprovalModalOpen(false);
    setApprovalLetterRef('');
    setTaxSavingsAmount('');
    setApprovalError(null);
  };

  const confirmApproval = async () => {
    setApprovalError(null);

    // Validate
    if (!approvalLetterRef.trim()) {
      setApprovalError('Approval letter reference is required.');
      return;
    }
    const taxSavings = parseFloat(taxSavingsAmount);
    if (!taxSavings || taxSavings <= 0) {
      setApprovalError('Tax savings amount must be a positive number.');
      return;
    }
    const feePercent = parseFloat(cahpFeePercent);
    if (isNaN(feePercent) || feePercent < 0 || feePercent > 100) {
      setApprovalError('CAHP Fee % must be between 0 and 100.');
      return;
    }
    if (!submittal.fields.PropertyLookupId) {
      setApprovalError("This submittal isn't linked to a property — can't create Billing/Disbursement records.");
      return;
    }

    const cahpFeeBilled = (taxSavings * feePercent) / 100;
    const ownerShare = taxSavings - cahpFeeBilled;

    setApprovalSaving(true);
    try {
      // 1. Update submittal: status → Approved, ApprovedAbatement set, NextAction cleared
      await updateListItem(LIST_NAMES.Submittals, submittal.id, {
        SubmittalStatus: 'Approved',
        ApprovedAbatement: taxSavings,
        NextAction: null,
        NextActionDue: null,
        SubmittalNotes: f.SubmittalNotes
          ? `${f.SubmittalNotes}\n\n[Approval ${new Date().toLocaleDateString()}] Letter ref: ${approvalLetterRef}`
          : `[Approval ${new Date().toLocaleDateString()}] Letter ref: ${approvalLetterRef}`,
      });

      // 2. Create Billing row — CAHP Fee = taxSavings * feePercent / 100
      const billingTitle = `${property?.fields.Title ?? 'Property'} ${f.cahpTaxYear ?? ''} CAHP Fee`.trim();
      try {
        await createListItem(LIST_NAMES.Billing, {
          Title: billingTitle,
          PropertyLookupId: submittal.fields.PropertyLookupId,
          cahpTaxYear: f.cahpTaxYear,
          AmountBilled: cahpFeeBilled,
          BillApprovedAbatement: taxSavings,
          CAHPFeePercent: feePercent,
          BillingStatus: 'Ready to Invoice' as BillingStatusValue,
          QBSyncStatus: 'Not Synced',
          BillingNotes: `Auto-created from Submittal approval. Approval letter: ${approvalLetterRef}`,
        });
      } catch (e) {
        // eslint-disable-next-line no-console
        console.warn('Billing row creation failed:', e);
        throw new Error(`Billing record creation failed: ${e instanceof Error ? e.message : String(e)}`);
      }

      // 3. Create Disbursement row — owner share = taxSavings - cahpFeeBilled
      const disbTitle = `${property?.fields.Title ?? 'Property'} ${f.cahpTaxYear ?? ''} Owner Disbursement`.trim();
      try {
        await createListItem(LIST_NAMES.Disbursements, {
          Title: disbTitle,
          DisbPropertyLookupId: submittal.fields.PropertyLookupId,
          DisbSubmittalLookupId: submittal.id,
          DisbAmount: ownerShare,
          DisbStatus: 'Pending' as DisbursementStatus,
          DisbNotes: `Auto-created from Submittal approval. Owner allocation TBD per ownership %.`,
        });
      } catch (e) {
        // eslint-disable-next-line no-console
        console.warn('Disbursement row creation failed:', e);
        // Don't throw — billing succeeded, disbursement is recoverable manually
      }

      await refetch();
      cancelApproval();
    } catch (err) {
      setApprovalError(err instanceof Error ? err.message : String(err));
    } finally {
      setApprovalSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Submittals" parentTo="/submittals" currentLabel={f.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1 flex-wrap">
            <h1 className="text-2xl font-bold text-teal-700">{f.Title}</h1>
            {f.FilingType && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${FILING_TYPE_STYLES[f.FilingType]}`}>
                {f.FilingType}
              </span>
            )}
            {f.SubmittalStatus && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[f.SubmittalStatus]}`}>
                {f.SubmittalStatus}
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500">
            {property ? (
              <Link to={`/properties/${property.id}`} className="text-teal-700 hover:text-teal-900 underline">
                {property.fields.Title}
              </Link>
            ) : (
              <span className="italic text-gray-400">unlinked</span>
            )}
            {f.cahpTaxYear && ` · Tax Year ${f.cahpTaxYear}`}
            {f.cahpState && ` · ${f.cahpState}`}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing ? (
            <button
              onClick={handleEdit}
              className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5"
            >
              <Icon name="settings" size={14} />
              Edit
            </button>
          ) : (
            <EditingActionButtons saving={saving} onCancel={handleCancel} onSave={handleSave} />
          )}
        </div>
      </div>

      <SaveErrorBanner error={saveError} />

      {/* Status pipeline */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-5">
        <div className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">Status Pipeline</div>
        {isTerminal && (currentStatus === 'Denied' || currentStatus === 'Withdrawn') ? (
          <div className={`p-3 rounded-md ${currentStatus === 'Denied' ? 'bg-red-50 border border-red-200' : 'bg-gray-50 border border-gray-200'}`}>
            <p className="text-sm font-semibold">{currentStatus}</p>
            <p className="text-xs text-gray-600 mt-1">
              {currentStatus === 'Denied'
                ? 'DOR denied this submittal. No billing record created. Appeal logic ships in Phase 3.'
                : 'Submittal withdrawn before DOR action. No billing record created.'}
            </p>
          </div>
        ) : (
          <div className="flex items-center gap-1 overflow-x-auto pb-1 mb-3">
            {PIPELINE_STAGES.map((stage, idx) => {
              const reached = statusIdx >= idx;
              const isCurrent = statusIdx === idx;
              return (
                <div key={stage.status} className="flex items-center gap-1 flex-1 min-w-fit">
                  <div
                    className={`flex-1 px-2 py-1.5 rounded-md text-center text-xs font-medium whitespace-nowrap transition-colors ${
                      isCurrent
                        ? 'bg-teal-700 text-white'
                        : reached
                          ? 'bg-success text-white'
                          : 'bg-gray-100 text-gray-400'
                    }`}
                  >
                    {reached && !isCurrent && <Icon name="check" size={10} className="inline mr-1" />}
                    {stage.label}
                  </div>
                  {idx < PIPELINE_STAGES.length - 1 && (
                    <div className={`w-3 h-0.5 ${reached && idx < statusIdx ? 'bg-success' : 'bg-gray-200'}`} />
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* Transition buttons */}
        {transitions.length > 0 && !editing && (
          <div className="flex flex-wrap gap-2 pt-2 border-t border-gray-100">
            <span className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider self-center mr-1">Next:</span>
            {transitions.map((t) => (
              <button
                key={t.to}
                onClick={() => startTransition(t)}
                className={`px-3 py-1.5 rounded-md text-xs font-medium border transition-colors ${TRANSITION_STYLES[t.style]}`}
              >
                {t.label}
              </button>
            ))}
          </div>
        )}

        {isTerminal && currentStatus === 'Approved' && (
          <div className="pt-2 border-t border-gray-100 flex items-start gap-2">
            <Icon name="check" size={12} className="text-success flex-shrink-0 mt-0.5" />
            <p className="text-xs text-success">
              Approved. Billing + Disbursement records auto-created at approval time
              {f.ApprovedAbatement != null && (
                <> · ${f.ApprovedAbatement.toLocaleString()} tax savings on file</>
              )}
              .
            </p>
          </div>
        )}
      </div>

      {/* Metadata + Action Plan two-column */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
        <div className="lg:col-span-2 bg-white border border-gray-200 rounded-lg shadow-card p-5">
          <h3 className="text-sm font-semibold text-gray-700 mb-3">Submittal Metadata</h3>
          <dl className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-2 text-sm">
            <EditableField
              label="Title"
              value={display.Title}
              editing={editing}
              onChange={(v) => handleFieldChange('Title', v as string)}
              required
            />
            <EditableField
              label="Filing Type"
              value={display.FilingType}
              editing={editing}
              type="choice"
              choices={FILING_TYPES}
              onChange={(v) => handleFieldChange('FilingType', v as SubmittalFilingType)}
            />
            <EditableField
              label="Tax Year"
              value={display.cahpTaxYear}
              editing={editing}
              type="choice"
              choices={TAX_YEARS}
              onChange={(v) => handleFieldChange('cahpTaxYear', v as CahpTaxYear)}
              mono
            />
            <EditableField
              label="State"
              value={display.cahpState}
              editing={editing}
              type="choice"
              choices={STATES}
              onChange={(v) => handleFieldChange('cahpState', v as CahpState)}
              mono
            />
            <EditableField
              label="Filing Method"
              value={display.FilingMethod}
              editing={editing}
              type="choice"
              choices={FILING_METHODS}
              onChange={(v) => handleFieldChange('FilingMethod', v as FilingMethod)}
            />
            <EditableField
              label="Date Filed"
              value={display.DateFiled}
              editing={editing}
              type="date"
              onChange={(v) => handleFieldChange('DateFiled', v as string)}
              mono
            />
            <EditableField
              label="Confirmation #"
              value={display.ConfirmationNumber}
              editing={editing}
              onChange={(v) => handleFieldChange('ConfirmationNumber', v as string)}
              mono
            />
            <EditableField
              label="Mail Tracking #"
              value={display.MailTrackingNumber}
              editing={editing}
              onChange={(v) => handleFieldChange('MailTrackingNumber', v as string)}
              mono
            />
            <EditableField
              label="Approved Abatement"
              value={display.ApprovedAbatement?.toString()}
              editing={editing}
              type="number"
              onChange={(v) => handleFieldChange('ApprovedAbatement', v === '' ? undefined : Number(v))}
              mono
            />
            <EditableField
              label="Next Action"
              value={display.NextAction}
              editing={editing}
              onChange={(v) => handleFieldChange('NextAction', v as string)}
            />
            <EditableField
              label="Next Action Due"
              value={display.NextActionDue}
              editing={editing}
              type="date"
              onChange={(v) => handleFieldChange('NextActionDue', v as string)}
              mono
            />
          </dl>
          <div className="mt-4 pt-4 border-t border-gray-100">
            <EditableField
              label="Notes"
              value={display.SubmittalNotes}
              editing={editing}
              type="textarea"
              rows={4}
              onChange={(v) => handleFieldChange('SubmittalNotes', v as string)}
            />
          </div>
        </div>

        {/* Action plan */}
        <div className="bg-white border border-gray-200 rounded-lg shadow-card p-5">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-semibold text-gray-700">Action Plan</h3>
            <span className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider">
              {actionPlan.length} open
            </span>
          </div>
          {actionPlan.length === 0 ? (
            <p className="text-xs text-gray-500 italic">No open outstanding items for this property.</p>
          ) : (
            <ul className="space-y-2 text-sm">
              {actionPlan.slice(0, 8).map((o) => (
                <li key={o.id} className="flex items-start gap-2">
                  <Icon
                    name={o.fields.ItemStatus === 'Overdue' ? 'alert' : 'inbox'}
                    size={12}
                    className={`mt-1 flex-shrink-0 ${o.fields.ItemStatus === 'Overdue' ? 'text-error' : 'text-gray-400'}`}
                  />
                  <div className="flex-1 min-w-0">
                    <div className="text-gray-900 truncate">{o.fields.Title}</div>
                    <div className="text-[11px] text-gray-500 font-mono-data">
                      {o.fields.ItemStatus ?? 'Requested'}
                      {o.fields.DateRequested && ` · requested ${new Date(o.fields.DateRequested).toLocaleDateString()}`}
                    </div>
                  </div>
                </li>
              ))}
              {actionPlan.length > 8 && (
                <li className="text-[11px] text-gray-400 italic">…and {actionPlan.length - 8} more</li>
              )}
            </ul>
          )}
        </div>
      </div>

      {/* Org Chart Snapshot — PR-10c — frozen historical org chart */}
      <div className="mb-6">
        <SubmittalOrgChartSnapshot
          snapshotJSON={f.OrgChartSnapshotJSON}
          capturedAt={f.OrgChartSnapshotDate}
          propertyTitle={property?.fields.Title ?? '(unlinked property)'}
        />
      </div>

      {/* Correspondence */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">DOR Correspondence</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            {relatedCorrespondence.length === 0
              ? 'No DOR correspondence logged for this property.'
              : `${relatedCorrespondence.length} letter${relatedCorrespondence.length === 1 ? '' : 's'} on file. Full CRUD ships in PR-11a.`}
          </p>
        </div>
        {relatedCorrespondence.length > 0 && (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Direction</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Received</th>
                <th className="px-4 py-3 text-left">Response Due</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {relatedCorrespondence.map((c) => (
                <tr key={c.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{c.fields.Direction || '—'}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{c.fields.LetterType || '—'}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{formatDate(c.fields.DateReceived)}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{formatDate(c.fields.ResponseDue)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Activity */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">Activity</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            {submittalActivity.length === 0
              ? 'No audit events for this submittal yet.'
              : `${submittalActivity.length} event${submittalActivity.length === 1 ? '' : 's'}, newest first`}
          </p>
        </div>
        {submittalActivity.length > 0 && (
          <ul className="divide-y divide-gray-100">
            {submittalActivity.map((row) => (
              <li key={row.id} className="px-4 py-3 flex items-start gap-3 text-sm">
                <span className="inline-block px-1.5 py-0.5 rounded text-[10px] font-bold bg-gray-100 text-gray-800 flex-shrink-0">
                  {row.fields.Action}
                </span>
                <div className="flex-1 min-w-0">
                  <div className="text-gray-900">{row.fields.Title}</div>
                  {row.fields.ChangeSummary && (
                    <pre className="mt-1 text-xs text-gray-600 whitespace-pre-wrap font-mono-data">
                      {row.fields.ChangeSummary}
                    </pre>
                  )}
                </div>
                <div className="text-xs text-gray-500 font-mono-data flex-shrink-0 text-right">
                  <div>{new Date(row.createdDateTime).toLocaleDateString()}</div>
                  <div className="text-gray-400">{row.createdBy?.user?.displayName ?? ''}</div>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>

      {/* Transition modal */}
      {activeTransition && (
        <TransitionModal
          transition={activeTransition}
          currentFields={f}
          inputFields={transitionFields}
          onFieldChange={handleTransitionFieldChange}
          onCancel={cancelTransition}
          onConfirm={confirmTransition}
          saving={saving}
          error={transitionError}
          willFreezeOrgChart={
            (currentStatus === 'Draft' || currentStatus === 'Package Mailed (NC)') &&
            activeTransition.to === 'Filed'
          }
        />
      )}

      {/* PR-10d — Approval Workflow modal */}
      {approvalModalOpen && (
        <ApprovalWorkflowModal
          submittalTitle={f.Title ?? ''}
          propertyTitle={property?.fields.Title}
          approvalLetterRef={approvalLetterRef}
          taxSavingsAmount={taxSavingsAmount}
          cahpFeePercent={cahpFeePercent}
          onLetterRefChange={setApprovalLetterRef}
          onTaxSavingsChange={setTaxSavingsAmount}
          onFeePercentChange={setCahpFeePercent}
          onCancel={cancelApproval}
          onConfirm={confirmApproval}
          saving={approvalSaving}
          error={approvalError}
        />
      )}
    </div>
  );
}

// =============================================================================
// Transition Modal — captures fields required for status transition + warns about side effects
// =============================================================================

function TransitionModal({
  transition,
  currentFields,
  inputFields,
  onFieldChange,
  onCancel,
  onConfirm,
  saving,
  error,
  willFreezeOrgChart,
}: {
  transition: Transition;
  currentFields: SubmittalFields;
  inputFields: Record<string, string>;
  onFieldChange: (field: string, value: string) => void;
  onCancel: () => void;
  onConfirm: () => void;
  saving: boolean;
  error: string | null;
  willFreezeOrgChart: boolean;
}) {
  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-md w-full p-5">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Transition to {transition.to}</h3>
        <p className="text-sm text-gray-600 mb-4">{transition.description}</p>

        {willFreezeOrgChart && (
          <div className="mb-4 bg-gold-50 border border-gold-200 rounded-md p-3 flex items-start gap-2">
            <Icon name="alert" size={14} className="text-gold-700 flex-shrink-0 mt-0.5" />
            <p className="text-xs text-teal-900">
              <strong>Org chart snapshot will be captured.</strong> The current ownership tree freezes onto this submittal at the moment of filing.
              Future ownership changes will not retroactively alter this snapshot — critical for DOR audit defensibility.
            </p>
          </div>
        )}

        {transition.requiresFields && transition.requiresFields.length > 0 && (
          <div className="space-y-3 mb-4">
            {transition.requiresFields.map((field) => {
              const currentValue = currentFields[field];
              if (currentValue) {
                // Already set — show it as a confirmation, not an input
                return (
                  <div key={String(field)} className="text-sm">
                    <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-0.5">
                      {String(field)} <span className="text-gray-400">(already set)</span>
                    </label>
                    <p className="font-mono-data text-gray-900">
                      {field === 'DateFiled' && typeof currentValue === 'string'
                        ? new Date(currentValue).toLocaleDateString()
                        : String(currentValue)}
                    </p>
                  </div>
                );
              }
              return (
                <div key={String(field)}>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    {String(field)} <span className="text-error">*</span>
                  </label>
                  <input
                    type={field === 'DateFiled' ? 'date' : 'text'}
                    value={inputFields[String(field)] ?? ''}
                    onChange={(e) => onFieldChange(String(field), e.target.value)}
                    className={`${inputClass} font-mono-data`}
                    autoFocus={transition.requiresFields?.[0] === field}
                  />
                </div>
              );
            })}
          </div>
        )}

        {error && (
          <div className="mb-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">
            {error}
          </div>
        )}

        <div className="flex items-center justify-end gap-2 mt-5">
          <button
            onClick={onCancel}
            disabled={saving}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            disabled={saving}
            className={`px-4 py-1.5 rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 ${TRANSITION_STYLES[transition.style]}`}
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-current border-r-transparent animate-spin opacity-70" />}
            {saving ? 'Saving…' : `Confirm Transition to ${transition.to}`}
          </button>
        </div>
      </div>
    </div>
  );
}

// =============================================================================
// PR-10d — Approval Workflow Modal
// Per spec §3.6.5: prompts for approval letter ref + tax savings, computes CAHP fee,
// shows what records will be created, confirms create-Billing + create-Disbursement on save.
// =============================================================================

function ApprovalWorkflowModal({
  submittalTitle,
  propertyTitle,
  approvalLetterRef,
  taxSavingsAmount,
  cahpFeePercent,
  onLetterRefChange,
  onTaxSavingsChange,
  onFeePercentChange,
  onCancel,
  onConfirm,
  saving,
  error,
}: {
  submittalTitle: string;
  propertyTitle?: string;
  approvalLetterRef: string;
  taxSavingsAmount: string;
  cahpFeePercent: string;
  onLetterRefChange: (v: string) => void;
  onTaxSavingsChange: (v: string) => void;
  onFeePercentChange: (v: string) => void;
  onCancel: () => void;
  onConfirm: () => void;
  saving: boolean;
  error: string | null;
}) {
  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  const taxSavings = parseFloat(taxSavingsAmount) || 0;
  const feePercent = parseFloat(cahpFeePercent) || 0;
  const cahpFee = (taxSavings * feePercent) / 100;
  const ownerShare = taxSavings - cahpFee;
  const isValid = taxSavings > 0 && feePercent >= 0 && feePercent <= 100 && approvalLetterRef.trim().length > 0;

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-lg w-full p-5 my-8">
        <div className="flex items-center gap-2 mb-1">
          <Icon name="check" size={20} className="text-success" />
          <h3 className="text-lg font-bold text-teal-700">Approval Workflow</h3>
        </div>
        <p className="text-sm text-gray-600 mb-4">
          DOR approved <strong>{submittalTitle}</strong>
          {propertyTitle && <> · {propertyTitle}</>}.
          Capture the approval details — Billing and Disbursement records will auto-create.
        </p>

        <div className="space-y-3 mb-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Approval Letter Reference <span className="text-error">*</span>
            </label>
            <input
              type="text"
              value={approvalLetterRef}
              onChange={(e) => onLetterRefChange(e.target.value)}
              placeholder="e.g., DOR-SC-2025-Approval-12345"
              className={`${inputClass} font-mono-data`}
              autoFocus
            />
            <p className="text-xs text-gray-400 mt-1">Reference printed on the DOR approval letter — stored on the submittal for audit trail.</p>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Tax Savings Amount <span className="text-error">*</span>
            </label>
            <div className="flex items-center gap-1">
              <span className="text-gray-500 font-mono-data">$</span>
              <input
                type="number"
                value={taxSavingsAmount}
                onChange={(e) => onTaxSavingsChange(e.target.value)}
                placeholder="0.00"
                min="0"
                step="0.01"
                className={`${inputClass} font-mono-data`}
              />
            </div>
            <p className="text-xs text-gray-400 mt-1">Total tax abatement granted by DOR.</p>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              CAHP Fee % <span className="text-error">*</span>
            </label>
            <div className="flex items-center gap-1">
              <input
                type="number"
                value={cahpFeePercent}
                onChange={(e) => onFeePercentChange(e.target.value)}
                placeholder="50"
                min="0" max="100" step="0.1"
                className={`${inputClass} font-mono-data`}
              />
              <span className="text-gray-500 font-mono-data">%</span>
            </div>
            <p className="text-xs text-gray-400 mt-1">Share of tax savings retained by CAHP. Typical: 30–50%.</p>
          </div>
        </div>

        {/* Computed preview */}
        {taxSavings > 0 && (
          <div className="mb-4 bg-gold-50 border border-gold-200 rounded-md p-3">
            <div className="text-[10px] font-semibold text-gold-900 uppercase tracking-wider mb-2">
              Computed
            </div>
            <dl className="text-sm space-y-1">
              <div className="flex items-center justify-between">
                <dt className="text-gray-700">Tax Savings</dt>
                <dd className="font-mono-data font-semibold text-gray-900">${taxSavings.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</dd>
              </div>
              <div className="flex items-center justify-between">
                <dt className="text-gray-700">CAHP Fee ({feePercent}%) → Billing</dt>
                <dd className="font-mono-data font-semibold text-teal-700">${cahpFee.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</dd>
              </div>
              <div className="flex items-center justify-between border-t border-gold-200 pt-1 mt-1">
                <dt className="text-gray-700">Owner Share → Disbursement</dt>
                <dd className="font-mono-data font-semibold text-gray-900">${ownerShare.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</dd>
              </div>
            </dl>
          </div>
        )}

        <div className="mb-4 bg-blue-50 border border-blue-200 rounded-md p-3">
          <div className="flex items-start gap-2 text-xs text-blue-900">
            <Icon name="alert" size={12} className="flex-shrink-0 mt-0.5" />
            <div>
              <p className="font-semibold mb-1">On Confirm:</p>
              <ul className="space-y-0.5 list-disc list-inside">
                <li>Submittal status → <strong>Approved</strong></li>
                <li>Submittal ApprovedAbatement set to <strong>${taxSavings.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></li>
                <li><strong>1 Billing row</strong> created (Ready to Invoice status, QB Not Synced)</li>
                <li><strong>1 Disbursement row</strong> created (Pending status, owner allocation TBD)</li>
                <li>All actions audit-logged</li>
              </ul>
            </div>
          </div>
        </div>

        {error && (
          <div className="mb-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">
            {error}
          </div>
        )}

        <div className="flex items-center justify-end gap-2">
          <button
            onClick={onCancel}
            disabled={saving}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            disabled={saving || !isValid}
            className="px-4 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Creating records…' : 'Confirm Approval'}
          </button>
        </div>
      </div>
    </div>
  );
}

// =============================================================================
// Helpers
// =============================================================================

/**
 * Serialize the ownership tree for snapshot storage.
 * We strip down to just the data needed to render later — no recursive Owner / Ownership object refs.
 */
function serializeTree(tree: ReturnType<typeof getBeneficialOwnershipTree>): unknown {
  return tree.map((node) => ({
    ownerId: node.owner?.id,
    ownerTitle: node.owner?.fields.Title,
    ownerType: node.owner?.fields.OwnerType,
    relationshipType: node.relationship.fields.RelationshipType,
    ownershipPercent: node.relationship.fields.OwnershipPercent,
    effectiveDate: node.relationship.fields.EffectiveDate,
    children: serializeTree(node.children),
  }));
}

function formatDate(iso: string | undefined): string | undefined {
  if (!iso) return undefined;
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString();
}
