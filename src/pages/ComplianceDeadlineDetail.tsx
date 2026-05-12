import { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  LIST_NAMES,
  type ComplianceDeadline,
  type ComplianceDeadlineFields,
  type DeadlineStatus,
  type Property,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
  formatDate,
} from '../components/detail';

const STATUS_STYLES: Record<DeadlineStatus, string> = {
  Upcoming: 'bg-blue-100 text-blue-800 border-blue-200',
  'In Progress': 'bg-amber-100 text-amber-800 border-amber-200',
  Completed: 'bg-green-100 text-green-800 border-green-200',
  Overdue: 'bg-red-100 text-red-800 border-red-200',
  Missed: 'bg-red-200 text-red-900 border-red-300',
};

const CHOICES = {
  DeadlineType: [
    'IRS 990 Filing',
    'Annual Recertification',
    'Rent Roll Review',
    'AMI Cert Renewal',
    'State Compliance Report',
    'Property Tax Filing',
    'Operating Agreement Review',
    'Other',
  ] as const,
  DeadlineStatus: ['Upcoming', 'In Progress', 'Completed', 'Overdue', 'Missed'] as const,
  Recurrence: ['One-Time', 'Annual', 'Quarterly', 'Monthly'] as const,
  AppliesTo: [
    'CAHP Entity',
    'All Properties',
    'Specific Property',
    'SC Portfolio',
    'NC Portfolio',
  ] as const,
  ResponsibleParty: ['Brandy', 'Chris', 'Brian', 'John', 'Aljon', 'Other'] as const,
  cahpState: ['SC', 'NC'] as const,
} as const;

export function ComplianceDeadlineDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<ComplianceDeadlineFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  const { data: deadline, loading, error, refetch } = useSharePointItem<ComplianceDeadline>(
    LIST_NAMES.ComplianceDeadlines,
    id
  );

  // Fetch properties so we can show the linked property name
  const { data: allProperties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 200 });
  const linkedProperty = deadline?.fields.PropertyLookupId
    ? allProperties?.find((p) => String(p.id) === String(deadline.fields.PropertyLookupId))
    : null;

  useEffect(() => {
    if (deadline && !editing) {
      setDraft({ ...deadline.fields });
    }
  }, [deadline?.id, deadline?.lastModifiedDateTime, editing]);

  const handleEdit = () => {
    if (!deadline) return;
    setDraft({ ...deadline.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    if (deadline) setDraft({ ...deadline.fields });
    setSaveError(null);
    setEditing(false);
  };

  const handleFieldChange = <K extends keyof ComplianceDeadlineFields>(
    field: K,
    value: ComplianceDeadlineFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleSave = async () => {
    if (!draft || !deadline) return;
    setSaving(true);
    setSaveError(null);
    try {
      const changes: Record<string, unknown> = {};
      (Object.keys(draft) as (keyof ComplianceDeadlineFields)[]).forEach((k) => {
        const oldVal = deadline.fields[k];
        const newVal = draft[k];
        if (oldVal !== newVal) {
          changes[k as string] = newVal === '' ? null : newVal;
        }
      });
      if (Object.keys(changes).length > 0) {
        await updateListItem(LIST_NAMES.ComplianceDeadlines, deadline.id, changes);
        await refetch();
      }
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  // Quick action: mark complete (sets status + completion date)
  const handleMarkComplete = async () => {
    if (!deadline) return;
    setSaving(true);
    setSaveError(null);
    try {
      await updateListItem(LIST_NAMES.ComplianceDeadlines, deadline.id, {
        DeadlineStatus: 'Completed',
        CompletionDate: new Date().toISOString(),
      });
      await refetch();
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
          <span className="text-sm">Loading deadline…</span>
        </div>
      </div>
    );
  }

  if (error || !deadline || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Compliance" parentTo="/compliance" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Deadline not found
          </div>
          <p className="text-sm text-red-700 mb-3">
            {error ? error.message : `No deadline with ID ${id} exists.`}
          </p>
          <button
            onClick={() => navigate('/compliance')}
            className="text-sm text-teal-700 hover:text-teal-900 font-medium underline"
          >
            ← Back to Compliance
          </button>
        </div>
      </div>
    );
  }

  const display = editing ? draft : deadline.fields;
  const isCompleted = deadline.fields.DeadlineStatus === 'Completed';
  const dueDate = deadline.fields.DueDate ? new Date(deadline.fields.DueDate) : null;
  const daysOut = dueDate ? Math.round((dueDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000)) : null;

  return (
    <div>
      <BreadcrumbBar
        parentLabel="Compliance"
        parentTo="/compliance"
        currentLabel={deadline.fields.Title}
      />

      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-3 flex-wrap">
            <h1 className="text-3xl font-bold text-teal-700">{deadline.fields.Title}</h1>
            {deadline.fields.DeadlineStatus && (
              <span
                className={`inline-block px-2.5 py-0.5 rounded text-xs font-semibold border ${
                  STATUS_STYLES[deadline.fields.DeadlineStatus]
                }`}
              >
                {deadline.fields.DeadlineStatus}
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500 mt-1">
            {[
              deadline.fields.DeadlineType,
              deadline.fields.ResponsibleParty,
              deadline.fields.cahpState,
              deadline.fields.Recurrence,
              dueDate
                ? daysOut !== null && daysOut < 0
                  ? `${Math.abs(daysOut)} days overdue`
                  : daysOut === 0
                    ? 'Due today'
                    : `Due in ${daysOut} day${daysOut === 1 ? '' : 's'}`
                : null,
            ]
              .filter(Boolean)
              .join(' · ')}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              {!isCompleted && (
                <button
                  onClick={handleMarkComplete}
                  disabled={saving}
                  className="px-3 py-1.5 bg-success hover:bg-green-700 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50"
                >
                  <Icon name="check" size={14} />
                  Mark Complete
                </button>
              )}
              <button
                onClick={handleEdit}
                className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors"
              >
                <Icon name="settings" size={14} />
                Edit
              </button>
            </>
          )}
          {editing && <EditingActionButtons saving={saving} onSave={handleSave} onCancel={handleCancel} />}
        </div>
      </div>

      <SaveErrorBanner error={saveError} />

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Section title="Deadline">
          <EditableField
            label="Description"
            value={display.Title}
            editing={editing}
            required
            onChange={(v) => handleFieldChange('Title', v as string)}
          />
          <EditableField
            label="Type"
            value={display.DeadlineType}
            editing={editing}
            type="choice"
            choices={CHOICES.DeadlineType}
            onChange={(v) => handleFieldChange('DeadlineType', v as ComplianceDeadlineFields['DeadlineType'])}
          />
          <EditableField
            label="Recurrence"
            value={display.Recurrence}
            editing={editing}
            type="choice"
            choices={CHOICES.Recurrence}
            onChange={(v) => handleFieldChange('Recurrence', v as ComplianceDeadlineFields['Recurrence'])}
          />
          <EditableField
            label="Applies To"
            value={display.AppliesTo}
            editing={editing}
            type="choice"
            choices={CHOICES.AppliesTo}
            onChange={(v) => handleFieldChange('AppliesTo', v as ComplianceDeadlineFields['AppliesTo'])}
          />
          {linkedProperty && (
            <div className="flex items-start gap-3">
              <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Linked Property</dt>
              <dd className="text-sm flex-1">
                <button
                  onClick={() => navigate(`/properties/${linkedProperty.id}`)}
                  className="text-teal-700 hover:text-teal-900 font-medium underline"
                >
                  {linkedProperty.fields.Title}
                </button>
              </dd>
            </div>
          )}
        </Section>

        <Section title="Status & Timing">
          <EditableField
            label="Status"
            value={display.DeadlineStatus}
            editing={editing}
            type="choice"
            choices={CHOICES.DeadlineStatus}
            onChange={(v) => handleFieldChange('DeadlineStatus', v as ComplianceDeadlineFields['DeadlineStatus'])}
          />
          <EditableField
            label="Due Date"
            value={display.DueDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('DueDate', v as string)}
          />
          <EditableField
            label="Completion Date"
            value={display.CompletionDate}
            editing={editing}
            type="date"
            onChange={(v) => handleFieldChange('CompletionDate', v as string)}
          />
        </Section>

        <Section title="Assignment">
          <EditableField
            label="Responsible Party"
            value={display.ResponsibleParty}
            editing={editing}
            type="choice"
            choices={CHOICES.ResponsibleParty}
            onChange={(v) => handleFieldChange('ResponsibleParty', v as ComplianceDeadlineFields['ResponsibleParty'])}
          />
          <EditableField
            label="State"
            value={display.cahpState}
            editing={editing}
            type="choice"
            choices={CHOICES.cahpState}
            mono
            onChange={(v) => handleFieldChange('cahpState', v as ComplianceDeadlineFields['cahpState'])}
          />
        </Section>

        <Section title="Audit Trail">
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Created</dt>
            <dd className="text-sm text-gray-900 flex-1 font-mono-data text-xs">
              {formatDate(deadline.createdDateTime)}
            </dd>
          </div>
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0">Last Modified</dt>
            <dd className="text-sm text-gray-900 flex-1 font-mono-data text-xs">
              {formatDate(deadline.lastModifiedDateTime)}
            </dd>
          </div>
        </Section>

        <Section title="Notes" fullWidth>
          <EditableField
            label="Notes"
            value={display.DeadlineNotes}
            editing={editing}
            type="textarea"
            rows={5}
            hideLabel
            onChange={(v) => handleFieldChange('DeadlineNotes', v as string)}
          />
        </Section>
      </div>
    </div>
  );
}
