import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Correspondence,
  type CorrespondenceFields,
  type Property,
  type Submittal,
  type LetterType,
  type CorrespondenceDirection,
  type CahpTaxYear,
  type CahpState,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';

const LETTER_TYPES: LetterType[] = [
  'Initial Acknowledgment',
  'Additional Info Request',
  'Org Chart Request',
  'Approval',
  'Denial',
  'Withdrawal Notice',
  'Refund Notice',
  'Other',
];

const DIRECTIONS: CorrespondenceDirection[] = ['Inbound (from DOR)', 'Outbound (to DOR)'];
const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];
const STATES: CahpState[] = ['SC', 'NC'];

export function CorrespondenceDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const { data: corr, loading, error, refetch } = useSharePointItem<Correspondence>(
    LIST_NAMES.Correspondence,
    id
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<CorrespondenceFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    if (corr && !editing) setDraft({ ...corr.fields });
  }, [corr?.id, corr?.lastModifiedDateTime, editing]);

  const property = useMemo(() => {
    if (!corr || !properties.data || !corr.fields.PropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(corr.fields.PropertyLookupId)) ?? null;
  }, [corr, properties.data]);

  const submittal = useMemo(() => {
    if (!corr || !submittals.data || !corr.fields.CorrSubmittalLookupId) return null;
    return submittals.data.find((s) => String(s.id) === String(corr.fields.CorrSubmittalLookupId)) ?? null;
  }, [corr, submittals.data]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading correspondence…</span>
        </div>
      </div>
    );
  }

  if (error || !corr || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Correspondence" parentTo="/correspondence" currentLabel="Letter Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load correspondence</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : corr.fields;

  const handleFieldChange = <K extends keyof CorrespondenceFields>(
    field: K,
    value: CorrespondenceFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...corr.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...corr.fields });
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
        const k = key as keyof CorrespondenceFields;
        if (draft[k] !== corr.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Correspondence, corr.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Permanently delete "${corr.fields.Title}"?\n\nThis only removes the correspondence record — it does NOT remove any Outstanding Items or revert Submittal status changes that were created via cascade. You'll need to clean those up manually if needed.`)) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Correspondence, corr.id);
      navigate('/correspondence');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Correspondence" parentTo="/correspondence" currentLabel={corr.fields.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">{corr.fields.Title}</h1>
          <p className="text-sm text-gray-500 mt-1">
            {property ? (
              <Link to={`/properties/${property.id}`} className="text-teal-700 hover:text-teal-900 underline">
                {property.fields.Title}
              </Link>
            ) : (
              <span className="italic text-gray-400">unlinked property</span>
            )}
            {submittal && (
              <>
                {' · '}
                <Link to={`/submittals/${submittal.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {submittal.fields.Title}
                </Link>
              </>
            )}
            {corr.fields.DateReceived && ` · ${new Date(corr.fields.DateReceived).toLocaleDateString()}`}
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
        <Section title="Letter Details">
          <EditableField
            label="Subject"
            value={display.Title}
            editing={editing}
            onChange={(v) => handleFieldChange('Title', v as string)}
            required
          />
          <EditableField
            label="Direction"
            value={display.Direction}
            editing={editing}
            type="choice"
            choices={DIRECTIONS}
            onChange={(v) => handleFieldChange('Direction', v as CorrespondenceDirection)}
          />
          <EditableField
            label="Letter Type"
            value={display.LetterType}
            editing={editing}
            type="choice"
            choices={LETTER_TYPES}
            onChange={(v) => handleFieldChange('LetterType', v as LetterType)}
          />
          <EditableField
            label="Date Received"
            value={display.DateReceived}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DateReceived', v as string)}
            mono
          />
          <EditableField
            label="Date Responded"
            value={display.DateResponded}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DateResponded', v as string)}
            mono
          />
          <EditableField
            label="Response Due"
            value={display.ResponseDue}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('ResponseDue', v as string)}
            mono
          />
        </Section>

        <Section title="Context">
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
          {/* Property and Submittal lookups are read-only at the detail level — change those via the parent record */}
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Property</dt>
            <dd className="text-sm flex-1">
              {property ? (
                <Link to={`/properties/${property.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {property.fields.Title}
                </Link>
              ) : <span className="text-gray-300">—</span>}
            </dd>
          </div>
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Submittal</dt>
            <dd className="text-sm flex-1">
              {submittal ? (
                <Link to={`/submittals/${submittal.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {submittal.fields.Title}
                </Link>
              ) : <span className="text-gray-300">—</span>}
            </dd>
          </div>
        </Section>

        <Section title="Summary & Response" fullWidth>
          <EditableField
            label="Request / Summary"
            value={display.RequestSummary}
            editing={editing}
            type="textarea"
            rows={4}
            onChange={(v) => handleFieldChange('RequestSummary', v as string)}
          />
          <EditableField
            label="Response Notes"
            value={display.ResponseNotes}
            editing={editing}
            type="textarea"
            rows={4}
            onChange={(v) => handleFieldChange('ResponseNotes', v as string)}
          />
        </Section>
      </div>
    </div>
  );
}
