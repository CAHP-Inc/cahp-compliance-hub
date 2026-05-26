import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  createListItem,
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
  type CorrChannel,
  type CorrespondencePropertyLink,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { formatDateET } from '../lib/dates';
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
const CHANNELS: CorrChannel[] = ['Letter', 'Email', 'Phone', 'Meeting', 'Other'];
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
  const propertyLinks = useSharePointList<CorrespondencePropertyLink>(LIST_NAMES.CorrespondencePropertyLinks, { top: 2000 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<CorrespondenceFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  // Multi-property linkage state
  const [selectedPropertyIds, setSelectedPropertyIds] = useState<Set<string>>(new Set());
  const [propertySearch, setPropertySearch] = useState('');

  const currentLinkedPropertyIds = useMemo(() => {
    if (!corr) return new Set<string>();
    const set = new Set<string>();
    (propertyLinks.data ?? []).forEach((l) => {
      if (String(l.fields.CorrLookupId ?? '') === String(corr.id) && l.fields.PropertyLookupId) {
        set.add(String(l.fields.PropertyLookupId));
      }
    });
    if (corr.fields.PropertyLookupId) set.add(String(corr.fields.PropertyLookupId));
    return set;
  }, [corr, propertyLinks.data]);

  useEffect(() => {
    if (corr && !editing) {
      setDraft({ ...corr.fields });
      setSelectedPropertyIds(new Set(currentLinkedPropertyIds));
    }
  }, [corr?.id, corr?.lastModifiedDateTime, editing, currentLinkedPropertyIds]);

  const linkedProperties = useMemo(() => {
    if (!properties.data) return [];
    return Array.from(currentLinkedPropertyIds)
      .map((id) => properties.data!.find((p) => String(p.id) === id))
      .filter((p): p is Property => !!p);
  }, [currentLinkedPropertyIds, properties.data]);

  const sortedProperties = useMemo(
    () => [...(properties.data ?? [])].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')),
    [properties.data],
  );
  const filteredEditProperties = useMemo(() => {
    const q = propertySearch.trim().toLowerCase();
    if (!q) return sortedProperties;
    return sortedProperties.filter((p) => (p.fields.Title ?? '').toLowerCase().includes(q));
  }, [sortedProperties, propertySearch]);

  const togglePropertyId = (id: string) =>
    setSelectedPropertyIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

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
      const propIds = Array.from(selectedPropertyIds);
      const primaryProp = propIds[0] ?? null;

      const changed: Record<string, unknown> = {};
      Object.keys(draft).forEach((key) => {
        const k = key as keyof CorrespondenceFields;
        if (draft[k] !== corr.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      // Sync legacy primary lookup with the first selected property
      if (String(corr.fields.PropertyLookupId ?? '') !== String(primaryProp ?? '')) {
        changed.PropertyLookupId = primaryProp;
      }

      if (Object.keys(changed).length > 0) {
        await updateListItem(LIST_NAMES.Correspondence, corr.id, changed);
      }

      // Diff junction rows
      const myLinks = (propertyLinks.data ?? []).filter(
        (l) => String(l.fields.CorrLookupId ?? '') === String(corr.id),
      );
      const existingIds = new Set(
        myLinks.map((l) => String(l.fields.PropertyLookupId ?? '')).filter(Boolean),
      );
      const toAdd = propIds.filter((id) => !existingIds.has(id));
      const toRemove = [...existingIds].filter((id) => !selectedPropertyIds.has(id));

      for (const pid of toAdd) {
        await createListItem(LIST_NAMES.CorrespondencePropertyLinks, {
          Title: `Corr ${corr.id} ↔ Property ${pid}`,
          CorrLookupId: Number(corr.id),
          PropertyLookupId: Number(pid),
        });
      }
      for (const pid of toRemove) {
        const row = myLinks.find((l) => String(l.fields.PropertyLookupId ?? '') === pid);
        if (row) await deleteListItem(LIST_NAMES.CorrespondencePropertyLinks, row.id);
      }

      await refetch();
      propertyLinks.refetch?.();
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
            {linkedProperties.length === 0 ? (
              <span className="italic text-gray-400">unlinked</span>
            ) : linkedProperties.length === 1 ? (
              <Link to={`/properties/${linkedProperties[0].id}`} className="text-teal-700 hover:text-teal-900 underline">
                {linkedProperties[0].fields.Title}
              </Link>
            ) : (
              <span>{linkedProperties.length} properties</span>
            )}
            {submittal && (
              <>
                {' · '}
                <Link to={`/submittals/${submittal.id}`} className="text-teal-700 hover:text-teal-900 underline">
                  {submittal.fields.Title}
                </Link>
              </>
            )}
            {corr.fields.DateReceived && ` · ${formatDateET(corr.fields.DateReceived)}`}
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
            label="Channel"
            value={display.CorrChannel ?? 'Letter'}
            editing={editing}
            type="choice"
            choices={CHANNELS}
            onChange={(v) => handleFieldChange('CorrChannel', v as CorrChannel)}
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
          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
              Properties{editing && <span className="block text-[10px] font-normal normal-case">{selectedPropertyIds.size} selected</span>}
            </dt>
            <dd className="text-sm flex-1">
              {!editing ? (
                linkedProperties.length === 0 ? (
                  <span className="text-gray-300">—</span>
                ) : (
                  <div className="flex flex-wrap gap-1">
                    {linkedProperties.map((p) => (
                      <Link
                        key={p.id}
                        to={`/properties/${p.id}`}
                        className="px-1.5 py-0.5 rounded bg-teal-50 text-teal-800 hover:bg-teal-100 text-[12px] font-medium"
                      >
                        {p.fields.Title}
                      </Link>
                    ))}
                  </div>
                )
              ) : (
                <div>
                  <input
                    type="text"
                    value={propertySearch}
                    onChange={(e) => setPropertySearch(e.target.value)}
                    placeholder="Search properties…"
                    disabled={saving}
                    className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500 mb-1"
                  />
                  <div className="border border-gray-300 rounded max-h-40 overflow-y-auto bg-white">
                    {filteredEditProperties.length === 0 ? (
                      <div className="px-2 py-2 text-[11px] text-gray-500 italic">No matches.</div>
                    ) : (
                      filteredEditProperties.map((p) => (
                        <label key={p.id} className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs">
                          <input
                            type="checkbox"
                            checked={selectedPropertyIds.has(String(p.id))}
                            onChange={() => togglePropertyId(String(p.id))}
                            disabled={saving}
                          />
                          <span className="flex-1 truncate">{p.fields.Title}</span>
                          {p.fields.cahpState && (
                            <span className="text-[10px] text-gray-500 flex-shrink-0">{p.fields.cahpState}</span>
                          )}
                        </label>
                      ))
                    )}
                  </div>
                  <p className="text-[11px] text-gray-500 mt-1">
                    Leave empty for portfolio-wide / general DOR comms.
                  </p>
                </div>
              )}
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
