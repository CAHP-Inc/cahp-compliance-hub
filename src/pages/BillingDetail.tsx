import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Billing,
  type BillingFields,
  type Property,
  type BillingStatusValue,
  type QBSyncStatus,
  type CahpTaxYear,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';

const STATUSES: BillingStatusValue[] = [
  'Pending Approval',
  'Ready to Invoice',
  'Invoiced',
  'Paid',
  'Disputed',
];
const QB_STATUSES: QBSyncStatus[] = ['Not Synced', 'Synced', 'Discrepancy'];
const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];

const STATUS_STYLES: Record<BillingStatusValue, string> = {
  'Pending Approval': 'bg-gray-100 text-gray-800',
  'Ready to Invoice': 'bg-amber-100 text-amber-800',
  'Invoiced': 'bg-blue-100 text-blue-800',
  'Paid': 'bg-green-100 text-green-800',
  'Disputed': 'bg-red-100 text-red-800',
};

export function BillingDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const { data: billing, loading, error, refetch } = useSharePointItem<Billing>(
    LIST_NAMES.Billing,
    id
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<BillingFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    if (billing && !editing) setDraft({ ...billing.fields });
  }, [billing?.id, billing?.lastModifiedDateTime, editing]);

  const property = useMemo(() => {
    if (!billing || !properties.data || !billing.fields.PropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(billing.fields.PropertyLookupId)) ?? null;
  }, [billing, properties.data]);

  // Computed: CAHP fee from tax savings + fee %
  const computedCAHPFee = useMemo(() => {
    if (!draft) return 0;
    const ts = draft.BillApprovedAbatement ?? 0;
    const pct = draft.CAHPFeePercent ?? 0;
    return (ts * pct) / 100;
  }, [draft]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading invoice…</span>
        </div>
      </div>
    );
  }

  if (error || !billing || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Billing" parentTo="/billing" currentLabel="Invoice Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load invoice</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : billing.fields;

  const handleFieldChange = <K extends keyof BillingFields>(field: K, value: BillingFields[K]) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...billing.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...billing.fields });
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
        const k = key as keyof BillingFields;
        if (draft[k] !== billing.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Billing, billing.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Permanently delete "${billing.fields.Title}"?\n\nThis cannot be undone.`)) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Billing, billing.id);
      navigate('/billing');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Billing" parentTo="/billing" currentLabel={billing.fields.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1 flex-wrap">
            <h1 className="text-2xl font-bold text-teal-700">{billing.fields.Title}</h1>
            {billing.fields.BillingStatus && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[billing.fields.BillingStatus]}`}>
                {billing.fields.BillingStatus}
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
            {billing.fields.cahpTaxYear && ` · Tax Year ${billing.fields.cahpTaxYear}`}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              <button
                onClick={handleDelete}
                disabled={saving}
                className="px-3 py-1.5 border border-red-300 text-error hover:bg-red-50 rounded-md text-sm font-medium disabled:opacity-50"
              >
                Delete
              </button>
              <button
                onClick={handleEdit}
                className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5"
              >
                <Icon name="settings" size={14} />
                Edit
              </button>
            </>
          )}
          {editing && <EditingActionButtons saving={saving} onCancel={handleCancel} onSave={handleSave} />}
        </div>
      </div>

      <SaveErrorBanner error={saveError} />

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Section title="Invoice Details">
          <EditableField
            label="Title"
            value={display.Title}
            editing={editing}
            onChange={(v) => handleFieldChange('Title', v as string)}
            required
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
            label="Status"
            value={display.BillingStatus}
            editing={editing}
            type="choice"
            choices={STATUSES}
            onChange={(v) => handleFieldChange('BillingStatus', v as BillingStatusValue)}
          />
          <EditableField
            label="Invoice Date"
            value={display.InvoiceDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('InvoiceDate', v as string)}
            mono
          />
          <EditableField
            label="Invoice #"
            value={display.InvoiceNumber}
            editing={editing}
            onChange={(v) => handleFieldChange('InvoiceNumber', v as string)}
            mono
          />
        </Section>

        <Section title="Financial">
          <EditableField
            label="Tax Savings (Abatement)"
            value={display.BillApprovedAbatement?.toString()}
            editing={editing}
            type="number"
            onChange={(v) => handleFieldChange('BillApprovedAbatement', v === '' ? undefined : Number(v))}
            mono
          />
          <EditableField
            label="CAHP Fee %"
            value={display.CAHPFeePercent?.toString()}
            editing={editing}
            type="number"
            onChange={(v) => handleFieldChange('CAHPFeePercent', v === '' ? undefined : Number(v))}
            mono
          />
          <EditableField
            label="Amount Billed"
            value={display.AmountBilled?.toString()}
            editing={editing}
            type="number"
            onChange={(v) => handleFieldChange('AmountBilled', v === '' ? undefined : Number(v))}
            mono
          />
          {editing && computedCAHPFee !== (display.AmountBilled ?? 0) && (
            <div className="flex items-start gap-3 py-1">
              <span className="text-xs text-gray-500 w-40 flex-shrink-0">Computed Fee</span>
              <div className="text-xs flex-1">
                <span className="font-mono-data text-teal-700">${computedCAHPFee.toFixed(2)}</span>
                <button
                  onClick={() => handleFieldChange('AmountBilled', computedCAHPFee)}
                  className="ml-2 text-[11px] text-teal-700 hover:text-teal-900 underline"
                >
                  use computed
                </button>
              </div>
            </div>
          )}
        </Section>

        <Section title="QuickBooks">
          <EditableField
            label="QB Sync Status"
            value={display.QBSyncStatus}
            editing={editing}
            type="choice"
            choices={QB_STATUSES}
            onChange={(v) => handleFieldChange('QBSyncStatus', v as QBSyncStatus)}
          />
          <div className="text-[11px] text-gray-500 italic mt-1">
            Manual workflow for now. When you push this invoice to QB, set the Invoice # above and flip QB Sync to <em>Synced</em>. Automated push deferred until QB API credentials are sorted.
          </div>
        </Section>

        <Section title="Notes" fullWidth>
          <EditableField
            label="Notes"
            value={display.BillingNotes}
            editing={editing}
            type="textarea"
            rows={5}
            hideLabel
            onChange={(v) => handleFieldChange('BillingNotes', v as string)}
          />
        </Section>
      </div>
    </div>
  );
}
