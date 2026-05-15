import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type OwnerCommunication,
  type OwnerCommunicationFields,
  type Property,
  type Owner,
  type CommType,
  type CommDirection,
  type CommStatus,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';

const COMM_TYPES: CommType[] = ['Email', 'Phone', 'Meeting', 'SMS', 'Other'];
const COMM_DIRECTIONS: CommDirection[] = ['Inbound', 'Outbound'];
const COMM_STATUSES: CommStatus[] = ['Open', 'Closed'];

const TYPE_STYLES: Record<CommType, string> = {
  Email: 'bg-blue-100 text-blue-800',
  Phone: 'bg-purple-100 text-purple-800',
  Meeting: 'bg-teal-100 text-teal-800',
  SMS: 'bg-amber-100 text-amber-800',
  Other: 'bg-gray-100 text-gray-700',
};

const STATUS_STYLES: Record<CommStatus, string> = {
  Open: 'bg-amber-100 text-amber-800',
  Closed: 'bg-gray-100 text-gray-500',
};

export function OwnerCommunicationDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const { data: comm, loading, error, refetch } = useSharePointItem<OwnerCommunication>(
    LIST_NAMES.Communications,
    id
  );
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<OwnerCommunicationFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    if (comm && !editing) setDraft({ ...comm.fields });
  }, [comm?.id, comm?.lastModifiedDateTime, editing]);

  const property = useMemo(() => {
    if (!comm || !properties.data || !comm.fields.CommPropertyLookupId) return null;
    return properties.data.find((p) => String(p.id) === String(comm.fields.CommPropertyLookupId)) ?? null;
  }, [comm, properties.data]);

  const owner = useMemo(() => {
    if (!comm || !owners.data || !comm.fields.CommOwnerLookupId) return null;
    return owners.data.find((o) => String(o.id) === String(comm.fields.CommOwnerLookupId)) ?? null;
  }, [comm, owners.data]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading communication…</span>
        </div>
      </div>
    );
  }

  if (error || !comm || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Owner Communications" parentTo="/comms" currentLabel="Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : comm.fields;

  const handleFieldChange = <K extends keyof OwnerCommunicationFields>(
    field: K,
    value: OwnerCommunicationFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...comm.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...comm.fields });
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
        const k = key as keyof OwnerCommunicationFields;
        if (draft[k] !== comm.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Communications, comm.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Permanently delete "${comm.fields.Title}"?\n\nThis cannot be undone. Outstanding Items spawned by this comm will remain.`)) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Communications, comm.id);
      navigate('/comms');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const handleQuickClose = async () => {
    setSaving(true);
    setSaveError(null);
    try {
      await updateListItem(LIST_NAMES.Communications, comm.id, { CommStatus: 'Closed' as CommStatus });
      await refetch();
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Owner Communications" parentTo="/comms" currentLabel={comm.fields.Title ?? ''} />

      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1 flex-wrap">
            <h1 className="text-2xl font-bold text-teal-700">{comm.fields.Title}</h1>
            {comm.fields.CommType && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${TYPE_STYLES[comm.fields.CommType]}`}>
                {comm.fields.CommDirection === 'Inbound' ? '← ' : comm.fields.CommDirection === 'Outbound' ? '→ ' : ''}
                {comm.fields.CommType}
              </span>
            )}
            {comm.fields.CommStatus && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[comm.fields.CommStatus]}`}>
                {comm.fields.CommStatus}
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500">
            {comm.fields.CommDate && new Date(comm.fields.CommDate).toLocaleDateString()}
            {property && (
              <>
                {' · '}
                <Link to={`/properties/${property.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {property.fields.Title}
                </Link>
              </>
            )}
            {owner && (
              <>
                {' · '}
                <Link to={`/owners/${owner.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {owner.fields.Title}
                </Link>
              </>
            )}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              {comm.fields.CommStatus === 'Open' && (
                <button
                  onClick={handleQuickClose}
                  disabled={saving}
                  className="px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50"
                >
                  <Icon name="check" size={14} />
                  Mark Closed
                </button>
              )}
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
        <Section title="Details">
          <EditableField
            label="Subject"
            value={display.Title}
            editing={editing}
            onChange={(v) => handleFieldChange('Title', v as string)}
            required
          />
          <EditableField
            label="Type"
            value={display.CommType}
            editing={editing}
            type="choice"
            choices={COMM_TYPES}
            onChange={(v) => handleFieldChange('CommType', v as CommType)}
          />
          <EditableField
            label="Direction"
            value={display.CommDirection}
            editing={editing}
            type="choice"
            choices={COMM_DIRECTIONS}
            onChange={(v) => handleFieldChange('CommDirection', v as CommDirection)}
          />
          <EditableField
            label="Status"
            value={display.CommStatus}
            editing={editing}
            type="choice"
            choices={COMM_STATUSES}
            onChange={(v) => handleFieldChange('CommStatus', v as CommStatus)}
          />
          <EditableField
            label="Date"
            value={display.CommDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('CommDate', v as string)}
            mono
          />
          <EditableField
            label="Response Due"
            value={display.CommResponseDue}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('CommResponseDue', v as string)}
            mono
          />
        </Section>

        <Section title="Participants & Links">
          <EditableField
            label="Participants"
            value={display.CommParticipants}
            editing={editing}
            onChange={(v) => handleFieldChange('CommParticipants', v as string)}
          />
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
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Owner</dt>
            <dd className="text-sm flex-1">
              {owner ? (
                <Link to={`/owners/${owner.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {owner.fields.Title}
                </Link>
              ) : <span className="text-gray-300">—</span>}
            </dd>
          </div>
        </Section>

        <Section title="Notes" fullWidth>
          <EditableField
            label="Notes"
            value={display.CommNotes}
            editing={editing}
            type="textarea"
            rows={6}
            hideLabel
            onChange={(v) => handleFieldChange('CommNotes', v as string)}
          />
        </Section>
      </div>
    </div>
  );
}
