import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Disbursement,
  type DisbursementFields,
  type Property,
  type Submittal,
  type Owner,
  type DisbursementStatus,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';

const STATUSES: DisbursementStatus[] = ['Pending', 'Issued', 'Cleared', 'Voided'];

const STATUS_STYLES: Record<DisbursementStatus, string> = {
  'Pending': 'bg-amber-100 text-amber-800',
  'Issued': 'bg-blue-100 text-blue-800',
  'Cleared': 'bg-green-100 text-green-800',
  'Voided': 'bg-gray-100 text-gray-500',
};

export function DisbursementDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const { data: disb, loading, error, refetch } = useSharePointItem<Disbursement>(
    LIST_NAMES.Disbursements,
    id
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<DisbursementFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    if (disb && !editing) setDraft({ ...disb.fields });
  }, [disb?.id, disb?.lastModifiedDateTime, editing]);

  const property = useMemo(() => {
    if (!disb || !properties.data || !disb.fields.DisbPropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(disb.fields.DisbPropertyLookupId)) ?? null;
  }, [disb, properties.data]);

  const submittal = useMemo(() => {
    if (!disb || !submittals.data || !disb.fields.DisbSubmittalLookupId) return null;
    return submittals.data.find((s) => String(s.id) === String(disb.fields.DisbSubmittalLookupId)) ?? null;
  }, [disb, submittals.data]);

  const owner = useMemo(() => {
    if (!disb || !owners.data || !disb.fields.DisbOwnerLookupId) return null;
    return owners.data.find((o) => String(o.id) === String(disb.fields.DisbOwnerLookupId)) ?? null;
  }, [disb, owners.data]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading disbursement…</span>
        </div>
      </div>
    );
  }

  if (error || !disb || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Billing" parentTo="/billing" currentLabel="Disbursement Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load disbursement</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : disb.fields;

  const handleFieldChange = <K extends keyof DisbursementFields>(field: K, value: DisbursementFields[K]) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...disb.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...disb.fields });
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
        const k = key as keyof DisbursementFields;
        if (draft[k] !== disb.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Disbursements, disb.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Permanently delete "${disb.fields.Title}"?\n\nThis cannot be undone.`)) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Disbursements, disb.id);
      navigate('/billing');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Billing" parentTo="/billing" currentLabel={disb.fields.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1 flex-wrap">
            <h1 className="text-2xl font-bold text-teal-700">{disb.fields.Title}</h1>
            {disb.fields.DisbStatus && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[disb.fields.DisbStatus]}`}>
                {disb.fields.DisbStatus}
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
            {submittal && (
              <>
                {' · '}
                <Link to={`/submittals/${submittal.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {submittal.fields.Title}
                </Link>
              </>
            )}
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
        <Section title="Disbursement Details">
          <EditableField
            label="Title"
            value={display.Title}
            editing={editing}
            onChange={(v) => handleFieldChange('Title', v as string)}
            required
          />
          <EditableField
            label="Status"
            value={display.DisbStatus}
            editing={editing}
            type="choice"
            choices={STATUSES}
            onChange={(v) => handleFieldChange('DisbStatus', v as DisbursementStatus)}
          />
          <EditableField
            label="Amount"
            value={display.DisbAmount?.toString()}
            editing={editing}
            type="number"
            onChange={(v) => handleFieldChange('DisbAmount', v === '' ? undefined : Number(v))}
            mono
          />
          <EditableField
            label="Check #"
            value={display.DisbCheckNum}
            editing={editing}
            onChange={(v) => handleFieldChange('DisbCheckNum', v as string)}
            mono
          />
        </Section>

        <Section title="Dates & Owner">
          <EditableField
            label="Issue Date"
            value={display.DisbIssueDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DisbIssueDate', v as string)}
            mono
          />
          <EditableField
            label="Clear Date"
            value={display.DisbClearDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DisbClearDate', v as string)}
            mono
          />
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Owner</dt>
            <dd className="text-sm flex-1">
              {owner ? (
                <Link to={`/owners/${owner.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {owner.fields.Title}
                </Link>
              ) : <span className="text-gray-300">— allocate via PR-12b Record Payment workflow</span>}
            </dd>
          </div>
        </Section>

        <Section title="Notes" fullWidth>
          <EditableField
            label="Notes"
            value={display.DisbNotes}
            editing={editing}
            type="textarea"
            rows={5}
            hideLabel
            onChange={(v) => handleFieldChange('DisbNotes', v as string)}
          />
        </Section>
      </div>
    </div>
  );
}
