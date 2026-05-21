import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  createListItem,
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
  type CommunicationPropertyLink,
  type CommunicationOwnerLink,
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
  const propertyLinks = useSharePointList<CommunicationPropertyLink>(LIST_NAMES.CommunicationPropertyLinks, { top: 2000 });
  const ownerLinks = useSharePointList<CommunicationOwnerLink>(LIST_NAMES.CommunicationOwnerLinks, { top: 2000 });

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<OwnerCommunicationFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  // Multi-linkage state — initialized from junction rows + legacy single fields
  const [selectedPropertyIds, setSelectedPropertyIds] = useState<Set<string>>(new Set());
  const [selectedOwnerIds, setSelectedOwnerIds] = useState<Set<string>>(new Set());
  const [propertySearch, setPropertySearch] = useState('');
  const [ownerSearch, setOwnerSearch] = useState('');

  // Build the current linked sets from junction + legacy primary
  const currentLinkedPropertyIds = useMemo(() => {
    if (!comm) return new Set<string>();
    const set = new Set<string>();
    (propertyLinks.data ?? []).forEach((l) => {
      if (String(l.fields.CommLookupId ?? '') === String(comm.id) && l.fields.PropertyLookupId) {
        set.add(String(l.fields.PropertyLookupId));
      }
    });
    if (comm.fields.CommPropertyLookupId) set.add(String(comm.fields.CommPropertyLookupId));
    return set;
  }, [comm, propertyLinks.data]);

  const currentLinkedOwnerIds = useMemo(() => {
    if (!comm) return new Set<string>();
    const set = new Set<string>();
    (ownerLinks.data ?? []).forEach((l) => {
      if (String(l.fields.CommLookupId ?? '') === String(comm.id) && l.fields.OwnerLookupId) {
        set.add(String(l.fields.OwnerLookupId));
      }
    });
    if (comm.fields.CommOwnerLookupId) set.add(String(comm.fields.CommOwnerLookupId));
    return set;
  }, [comm, ownerLinks.data]);

  useEffect(() => {
    if (comm && !editing) {
      setDraft({ ...comm.fields });
      setSelectedPropertyIds(new Set(currentLinkedPropertyIds));
      setSelectedOwnerIds(new Set(currentLinkedOwnerIds));
    }
  }, [comm?.id, comm?.lastModifiedDateTime, editing, currentLinkedPropertyIds, currentLinkedOwnerIds]);

  const linkedProperties = useMemo(() => {
    if (!properties.data) return [];
    return Array.from(currentLinkedPropertyIds)
      .map((id) => properties.data!.find((p) => String(p.id) === id))
      .filter((p): p is Property => !!p);
  }, [currentLinkedPropertyIds, properties.data]);

  const linkedOwners = useMemo(() => {
    if (!owners.data) return [];
    return Array.from(currentLinkedOwnerIds)
      .map((id) => owners.data!.find((o) => String(o.id) === id))
      .filter((o): o is Owner => !!o);
  }, [currentLinkedOwnerIds, owners.data]);

  const sortedProperties = useMemo(
    () =>
      [...(properties.data ?? [])].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')),
    [properties.data],
  );
  const sortedOwners = useMemo(
    () =>
      [...(owners.data ?? [])].sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '')),
    [owners.data],
  );
  const filteredEditProperties = useMemo(() => {
    const q = propertySearch.trim().toLowerCase();
    if (!q) return sortedProperties;
    return sortedProperties.filter((p) => (p.fields.Title ?? '').toLowerCase().includes(q));
  }, [sortedProperties, propertySearch]);
  const filteredEditOwners = useMemo(() => {
    const q = ownerSearch.trim().toLowerCase();
    if (!q) return sortedOwners;
    return sortedOwners.filter((o) => (o.fields.Title ?? '').toLowerCase().includes(q));
  }, [sortedOwners, ownerSearch]);

  const togglePropertyId = (id: string) =>
    setSelectedPropertyIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const toggleOwnerId = (id: string) =>
    setSelectedOwnerIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

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
      // Keep the legacy primary lookup in sync with the first selected linkage,
      // so SharePoint default views still surface a meaningful value.
      const propIds = Array.from(selectedPropertyIds);
      const ownerIds = Array.from(selectedOwnerIds);
      const primaryProp = propIds[0] ?? null;
      const primaryOwner = ownerIds[0] ?? null;

      const changed: Record<string, unknown> = {};
      Object.keys(draft).forEach((key) => {
        const k = key as keyof OwnerCommunicationFields;
        if (draft[k] !== comm.fields[k]) {
          changed[k] = draft[k] === '' ? null : draft[k];
        }
      });
      // Force the primary lookup columns to reflect the new selection
      if (String(comm.fields.CommPropertyLookupId ?? '') !== String(primaryProp ?? '')) {
        changed.CommPropertyLookupId = primaryProp;
      }
      if (String(comm.fields.CommOwnerLookupId ?? '') !== String(primaryOwner ?? '')) {
        changed.CommOwnerLookupId = primaryOwner;
      }

      if (Object.keys(changed).length > 0) {
        await updateListItem(LIST_NAMES.Communications, comm.id, changed);
      }

      // Diff junction rows
      const myPropertyLinks = (propertyLinks.data ?? []).filter(
        (l) => String(l.fields.CommLookupId ?? '') === String(comm.id),
      );
      const myOwnerLinks = (ownerLinks.data ?? []).filter(
        (l) => String(l.fields.CommLookupId ?? '') === String(comm.id),
      );
      const existingPropIds = new Set(
        myPropertyLinks.map((l) => String(l.fields.PropertyLookupId ?? '')).filter(Boolean),
      );
      const existingOwnerIds = new Set(
        myOwnerLinks.map((l) => String(l.fields.OwnerLookupId ?? '')).filter(Boolean),
      );

      const propsToAdd = propIds.filter((id) => !existingPropIds.has(id));
      const propsToRemove = [...existingPropIds].filter((id) => !selectedPropertyIds.has(id));
      const ownersToAdd = ownerIds.filter((id) => !existingOwnerIds.has(id));
      const ownersToRemove = [...existingOwnerIds].filter((id) => !selectedOwnerIds.has(id));

      for (const pid of propsToAdd) {
        await createListItem(LIST_NAMES.CommunicationPropertyLinks, {
          Title: `Comm ${comm.id} ↔ Property ${pid}`,
          CommLookupId: Number(comm.id),
          PropertyLookupId: Number(pid),
        });
      }
      for (const pid of propsToRemove) {
        const row = myPropertyLinks.find((l) => String(l.fields.PropertyLookupId ?? '') === pid);
        if (row) await deleteListItem(LIST_NAMES.CommunicationPropertyLinks, row.id);
      }
      for (const oid of ownersToAdd) {
        await createListItem(LIST_NAMES.CommunicationOwnerLinks, {
          Title: `Comm ${comm.id} ↔ Owner ${oid}`,
          CommLookupId: Number(comm.id),
          OwnerLookupId: Number(oid),
        });
      }
      for (const oid of ownersToRemove) {
        const row = myOwnerLinks.find((l) => String(l.fields.OwnerLookupId ?? '') === oid);
        if (row) await deleteListItem(LIST_NAMES.CommunicationOwnerLinks, row.id);
      }

      await refetch();
      propertyLinks.refetch?.();
      ownerLinks.refetch?.();
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
            {linkedProperties.length > 0 && (
              <>
                {' · '}
                {linkedProperties.length === 1 ? 'property' : `${linkedProperties.length} properties`}
              </>
            )}
            {linkedOwners.length > 0 && (
              <>
                {' · '}
                {linkedOwners.length === 1 ? 'owner' : `${linkedOwners.length} owners`}
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
                </div>
              )}
            </dd>
          </div>

          <div className="flex items-start gap-3">
            <dt className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
              Owner Entities{editing && <span className="block text-[10px] font-normal normal-case">{selectedOwnerIds.size} selected</span>}
            </dt>
            <dd className="text-sm flex-1">
              {!editing ? (
                linkedOwners.length === 0 ? (
                  <span className="text-gray-300">—</span>
                ) : (
                  <div className="flex flex-wrap gap-1">
                    {linkedOwners.map((o) => (
                      <Link
                        key={o.id}
                        to={`/owners/${o.id}`}
                        className="px-1.5 py-0.5 rounded bg-gold-50 text-gold-900 hover:bg-gold-100 text-[12px] font-medium"
                      >
                        {o.fields.Title}
                      </Link>
                    ))}
                  </div>
                )
              ) : (
                <div>
                  <input
                    type="text"
                    value={ownerSearch}
                    onChange={(e) => setOwnerSearch(e.target.value)}
                    placeholder="Search owner entities…"
                    disabled={saving}
                    className="w-full px-2 py-1 border border-gray-300 rounded text-xs focus:outline-none focus:border-teal-500 mb-1"
                  />
                  <div className="border border-gray-300 rounded max-h-40 overflow-y-auto bg-white">
                    {filteredEditOwners.length === 0 ? (
                      <div className="px-2 py-2 text-[11px] text-gray-500 italic">No matches.</div>
                    ) : (
                      filteredEditOwners.map((o) => (
                        <label key={o.id} className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs">
                          <input
                            type="checkbox"
                            checked={selectedOwnerIds.has(String(o.id))}
                            onChange={() => toggleOwnerId(String(o.id))}
                            disabled={saving}
                          />
                          <span className="flex-1 truncate">{o.fields.Title}</span>
                          {o.fields.OwnerType && (
                            <span className="text-[10px] text-gray-500 flex-shrink-0">{o.fields.OwnerType}</span>
                          )}
                        </label>
                      ))
                    )}
                  </div>
                </div>
              )}
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
