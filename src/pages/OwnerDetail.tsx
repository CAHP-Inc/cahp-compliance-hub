import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  getDirectOwnersOf,
  countPropertiesForOwner,
  countMembersOf,
  countLLCsOwnedBy,
  getPropertyIdsForOwner,
  getDownstreamTree,
  getAncestorOwnerIds,
  type DownstreamNode,
  type Owner,
  type OwnerFields,
  type Ownership,
  type Property,
  type OwnerType,
  type OutstandingItem,
  type Contact,
  type ContactOwnerLink,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { EntityDocumentsSection } from '../components/EntityDocumentsSection';
import { LoggedCommunicationsSection } from '../components/LoggedCommunicationsSection';
import { DeedsSection } from '../components/DeedsSection';
import { EditOwnershipModal } from '../components/EditOwnershipModal';
import { ExportOutstandingItemsModal } from '../components/ExportOutstandingItemsModal';
import { formatDateOnly } from '../lib/dates';
import {
  BreadcrumbBar,
  Section,
  EditableField,
  SaveErrorBanner,
  EditingActionButtons,
} from '../components/detail';

const TYPE_STYLES: Record<OwnerType, string> = {
  Individual: 'bg-blue-100 text-blue-800',
  LLC: 'bg-purple-100 text-purple-800',
  Nonprofit: 'bg-teal-100 text-teal-800',
  Trust: 'bg-amber-100 text-amber-800',
  Corporation: 'bg-indigo-100 text-indigo-800',
  'Limited Partnership': 'bg-rose-100 text-rose-800',
  'General Partnership': 'bg-fuchsia-100 text-fuchsia-800',
};

const TYPE_OPTIONS: OwnerType[] = ['Individual', 'LLC', 'Nonprofit', 'Trust', 'Corporation', 'Limited Partnership', 'General Partnership'];

export function OwnerDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();

  const { data: owner, loading, error, refetch } = useSharePointItem<Owner>(LIST_NAMES.Owners, id);
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const [editing, setEditing] = useState(false);
  const [editingOwnershipId, setEditingOwnershipId] = useState<string | null>(null);
  const [draft, setDraft] = useState<OwnerFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    if (owner && !editing) setDraft({ ...owner.fields });
  }, [owner?.id, owner?.lastModifiedDateTime, editing]);

  const directProperties = useMemo(() => {
    if (!ownership.data || !properties.data || !id) return [];
    const propsById = new Map(properties.data.map((p) => [String(p.id), p]));
    return ownership.data
      .filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(id) &&
          o.fields.LinkedPropertyLookupId
      )
      .map((o) => ({
        relationship: o,
        property: propsById.get(String(o.fields.LinkedPropertyLookupId)) ?? null,
      }))
      .filter((x) => x.property);
  }, [ownership.data, properties.data, id]);

  const directMemberships = useMemo(() => {
    if (!ownership.data || !owners.data || !id) return [];
    const ownersById = new Map(owners.data.map((o) => [String(o.id), o]));
    return ownership.data
      .filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(id) &&
          o.fields.ParentOwnerLookupId
      )
      .map((o) => ({
        relationship: o,
        parent: ownersById.get(String(o.fields.ParentOwnerLookupId)) ?? null,
      }))
      .filter((x) => x.parent);
  }, [ownership.data, owners.data, id]);

  const members = useMemo(() => {
    if (!ownership.data || !owners.data || !id) return [];
    return getDirectOwnersOf('owner', id, ownership.data, owners.data);
  }, [ownership.data, owners.data, id]);

  const stats = useMemo(() => {
    if (!ownership.data || !owners.data || !id) return null;
    return {
      propertyCounts: countPropertiesForOwner(id, ownership.data, owners.data),
      memberCount: countMembersOf(id, ownership.data),
      llcsOwned: countLLCsOwnedBy(id, ownership.data),
    };
  }, [ownership.data, owners.data, id]);

  if (loading) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading owner…</span>
        </div>
      </div>
    );
  }

  if (error || !owner || !draft) {
    return (
      <div>
        <BreadcrumbBar parentLabel="Owners" parentTo="/owners" currentLabel="Owner Detail" />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load owner</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error?.message ?? 'Not found'}</p>
        </div>
      </div>
    );
  }

  const display = editing ? draft : owner.fields;

  const handleFieldChange = <K extends keyof OwnerFields>(field: K, value: OwnerFields[K]) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleEdit = () => {
    setDraft({ ...owner.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    setDraft({ ...owner.fields });
    setSaveError(null);
    setEditing(false);
  };

  const handleSave = async () => {
    if (!draft) return;
    setSaving(true);
    setSaveError(null);
    try {
      // Compute diff
      const changed: Record<string, unknown> = {};
      Object.keys(draft).forEach((key) => {
        const k = key as keyof OwnerFields;
        if (draft[k] !== owner.fields[k]) changed[k] = draft[k];
      });
      if (Object.keys(changed).length === 0) {
        setEditing(false);
        return;
      }
      await updateListItem(LIST_NAMES.Owners, owner.id, changed);
      await refetch();
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    const totalRelations = (stats?.memberCount ?? 0) + directProperties.length + directMemberships.length;
    if (totalRelations > 0) {
      alert(
        `Can't delete — this owner appears in ${totalRelations} ownership relationship${totalRelations === 1 ? '' : 's'}. ` +
          `Remove those records first.`
      );
      return;
    }
    if (!confirm(`Permanently delete "${owner.fields.Title}"? This cannot be undone.`)) return;
    setSaving(true);
    try {
      await deleteListItem(LIST_NAMES.Owners, owner.id);
      navigate('/owners');
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  return (
    <div>
      <BreadcrumbBar parentLabel="Owners" parentTo="/owners" currentLabel={owner.fields.Title ?? ''} />

      {/* Header */}
      <div className="flex items-start justify-between mb-6">
        <div>
          <div className="flex items-center gap-3 mb-1">
            <h1 className="text-2xl font-bold text-teal-700">{owner.fields.Title}</h1>
            {owner.fields.OwnerType && (
              <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${TYPE_STYLES[owner.fields.OwnerType]}`}>
                {owner.fields.OwnerType}
              </span>
            )}
          </div>
          {owner.fields.OwnerState && (
            <p className="text-sm text-gray-500">
              {owner.fields.OwnerType === 'Individual' ? 'Resides in' : 'Formed in'} {owner.fields.OwnerState}
            </p>
          )}
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

      {/* KPI Row */}
      {stats && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <KPI label="Direct Properties" value={stats.propertyCounts.direct} />
          <KPI label="Indirect Properties" value={stats.propertyCounts.indirect} sub="via owned LLCs" />
          <KPI label="Members" value={stats.memberCount} sub="entities with stake here" />
          <KPI label="LLCs Owned" value={stats.llcsOwned} sub="entities this owner is in" />
        </div>
      )}

      {/* Identity */}
      <Section title="Identity">
        <EditableField
          label="Legal Name"
          value={display.Title}
          editing={editing}
          onChange={(v) => handleFieldChange('Title', v as string)}
          required
        />
        <EditableField
          label="Owner Type"
          value={display.OwnerType}
          editing={editing}
          type="choice"
          choices={TYPE_OPTIONS}
          onChange={(v) => handleFieldChange('OwnerType', v as OwnerType)}
        />
        <EditableField
          label="State"
          value={display.OwnerState}
          editing={editing}
          onChange={(v) => handleFieldChange('OwnerState', v as string)}
        />
        <EditableField
          label="Tax ID"
          value={editing ? display.TaxID : maskTaxID(display.TaxID)}
          editing={editing}
          onChange={(v) => handleFieldChange('TaxID', v as string)}
          mono
        />
        <EditableField
          label="Contact Email"
          value={display.ContactEmail}
          editing={editing}
          onChange={(v) => handleFieldChange('ContactEmail', v as string)}
        />
        <EditableField
          label="Sponsor / Principal"
          value={display.SponsorName}
          editing={editing}
          onChange={(v) => handleFieldChange('SponsorName', v as string)}
        />
        <EditableField
          label="Entity Description (override)"
          value={display.EntityDescription}
          editing={editing}
          onChange={(v) => handleFieldChange('EntityDescription', v as string)}
        />
        <EditableField
          label="Address"
          value={display.OwnerAddress}
          editing={editing}
          type="textarea"
          rows={2}
          onChange={(v) => handleFieldChange('OwnerAddress', v as string)}
        />
        {(owner.fields.OwnerType === 'Nonprofit' || display.OwnerType === 'Nonprofit') && (
          <EditableField
            label="Tax Exempt (501(c)(3))"
            value={display.IsTaxExempt ? 'Yes' : 'No'}
            editing={editing}
            type="choice"
            choices={['Yes', 'No'] as const}
            onChange={(v) => handleFieldChange('IsTaxExempt', v === 'Yes' as never)}
          />
        )}
        <EditableField
          label="CAHP Entity"
          value={display.IsCAHPEntity ? 'Yes' : 'No'}
          editing={editing}
          type="choice"
          choices={['Yes', 'No'] as const}
          onChange={(v) => handleFieldChange('IsCAHPEntity', v === 'Yes' as never)}
        />
        {(display.IsCAHPEntity || owner.fields.IsCAHPEntity) && (
          <p className="text-[11px] text-gold-900 italic ml-44">
            Flagged as part of the CAHP family — appears in gold on org charts. Subsidiary LLCs with this entity as a member are flagged as "exemption source" entities (their docs accompany DOR filings).
          </p>
        )}
        <EditableField
          label="Notes"
          value={display.OwnerNotes}
          editing={editing}
          type="textarea"
          onChange={(v) => handleFieldChange('OwnerNotes', v as string)}
        />
      </Section>

      {/* Members section — for entities that can have upstream owners (LLC, Nonprofit, Trust, Corp, Partnership) */}
      {(owner.fields.OwnerType && owner.fields.OwnerType !== 'Individual') && (
        <div className="bg-white border-l-4 border-gold-500 border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
          <div className="px-4 py-3 border-b border-gray-100 bg-gold-50 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-semibold text-teal-900">Members of {owner.fields.Title}</h3>
              <p className="text-xs text-gray-600 mt-0.5">
                Edits cascade to every property where this entity holds interest.
                {stats && stats.propertyCounts.direct + stats.propertyCounts.indirect > 0 &&
                  ` Currently ${stats.propertyCounts.direct + stats.propertyCounts.indirect} property${(stats.propertyCounts.direct + stats.propertyCounts.indirect) === 1 ? '' : 'ies'} affected.`}
              </p>
            </div>
            <button
              onClick={() => navigate(`/ownership/new?parentOwnerId=${owner.id}`)}
              className="text-xs text-teal-700 hover:text-teal-900 font-medium flex items-center gap-1 flex-shrink-0"
            >
              <Icon name="plus" size={12} />
              Add Member
            </button>
          </div>
          {members.length === 0 ? (
            <div className="px-4 py-8 text-center text-sm text-gray-500">
              No members on file. Add the first member with the link above.
            </div>
          ) : (
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
                <tr>
                  <th className="px-4 py-3 text-left">Entity</th>
                  <th className="px-4 py-3 text-left">Type</th>
                  <th className="px-4 py-3 text-left">Role</th>
                  <th className="px-4 py-3 text-right">%</th>
                  <th className="px-4 py-3 text-left">Effective</th>
                  <th className="px-4 py-3 text-right w-16"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {members.map(({ relationship: rel, owner: memberOwner }) => (
                  <tr
                    key={rel.id}
                    onClick={() => memberOwner && navigate(`/owners/${memberOwner.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 font-medium text-gray-900">
                      {memberOwner?.fields.Title ?? '(unresolved)'}
                    </td>
                    <td className="px-4 py-3">
                      {memberOwner?.fields.OwnerType && (
                        <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold ${TYPE_STYLES[memberOwner.fields.OwnerType]}`}>
                          {memberOwner.fields.OwnerType}
                        </span>
                      )}
                    </td>
                    <td className="px-4 py-3 text-xs text-gray-700">{rel.fields.RelationshipType ?? '—'}</td>
                    <td className="px-4 py-3 text-right font-mono-data">
                      {rel.fields.OwnershipPercent != null ? `${rel.fields.OwnershipPercent}%` : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                      {rel.fields.EffectiveDate ? new Date(rel.fields.EffectiveDate).toLocaleDateString() : '—'}
                    </td>
                    <td className="px-4 py-3 text-right">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setEditingOwnershipId(rel.id);
                        }}
                        className="text-[11px] text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50 transition-colors"
                        title="Edit all fields on this membership record"
                      >
                        Edit
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}

      {/* Cascade Preview — list of properties affected by member changes (any entity with potential members) */}
      {(owner.fields.OwnerType && owner.fields.OwnerType !== 'Individual') && (
        <CascadePreviewSection ownerId={owner.id} ownership={ownership.data ?? []} owners={owners.data ?? []} />
      )}

      {/* Owns Interest In (other LLCs) */}
      {directMemberships.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
          <div className="px-4 py-3 border-b border-gray-100">
            <h3 className="text-sm font-semibold text-gray-700">Owns Interest In</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {directMemberships.length} {directMemberships.length === 1 ? 'entity' : 'entities'} where this owner is a member
            </p>
          </div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Entity</th>
                <th className="px-4 py-3 text-left">Role</th>
                <th className="px-4 py-3 text-right">%</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {directMemberships.map(({ relationship: rel, parent }) => (
                <tr
                  key={rel.id}
                  onClick={() => parent && navigate(`/owners/${parent.id}`)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">{parent?.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{rel.fields.RelationshipType ?? '—'}</td>
                  <td className="px-4 py-3 text-right font-mono-data">
                    {rel.fields.OwnershipPercent != null ? `${rel.fields.OwnershipPercent}%` : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Direct Property Holdings */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100 flex items-start justify-between gap-3">
          <div>
            <h3 className="text-sm font-semibold text-gray-700">Property Holdings</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {directProperties.length} direct property {directProperties.length === 1 ? 'holding' : 'holdings'}
              {' · '}
              <span className="text-gray-400">Add properties directly owned by this entity to track their tax map IDs and filings.</span>
            </p>
          </div>
          <button
            onClick={() => navigate(`/properties/new?ownerId=${owner.id}&role=Member&percent=100`)}
            className="text-xs text-teal-700 hover:text-teal-900 font-medium flex items-center gap-1 flex-shrink-0"
            title={`Create a new property pre-linked to ${owner.fields.Title}`}
          >
            <Icon name="plus" size={12} />
            Add Property
          </button>
        </div>
        {directProperties.length === 0 ? (
          <div className="px-4 py-8 text-center text-sm text-gray-500">
            This owner doesn't directly hold any properties.
            {stats && stats.propertyCounts.indirect > 0 && (
              <> Has indirect interest in <strong>{stats.propertyCounts.indirect}</strong> propert{stats.propertyCounts.indirect === 1 ? 'y' : 'ies'} via owned entities.</>
            )}
            <div className="mt-3 text-xs text-gray-400">
              Use <strong>Add Property</strong> above to create a property record for an SFR or portfolio this entity owns outright — then add tax map IDs to it from the property page.
            </div>
          </div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">Role</th>
                <th className="px-4 py-3 text-right">%</th>
                <th className="px-4 py-3 text-left">Effective</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {directProperties.map(({ relationship: rel, property }) => (
                <tr
                  key={rel.id}
                  onClick={() => property && navigate(`/properties/${property.id}`)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">{property?.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{rel.fields.RelationshipType ?? '—'}</td>
                  <td className="px-4 py-3 text-right font-mono-data">
                    {rel.fields.OwnershipPercent != null ? `${rel.fields.OwnershipPercent}%` : '—'}
                  </td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                    {rel.fields.EffectiveDate ? new Date(rel.fields.EffectiveDate).toLocaleDateString() : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Downstream tree — everything this entity owns or holds interest in */}
      <OwnerDownstreamTreeSection
        ownerId={String(owner.id)}
        ownerTitle={owner.fields.Title}
        ownership={ownership.data ?? []}
        owners={owners.data ?? []}
        properties={properties.data ?? []}
      />

      {/* Emails + manually-logged communications tied to this owner */}
      <LoggedCommunicationsSection
        ownerId={String(owner.id)}
        title="Email Log"
        subtitle="Emails sent to this owner's contacts via the Compose modal auto-log here, plus any manual log entries."
      />

      {/* Outstanding items across every property this owner has an interest in */}
      <OwnerOutstandingItemsSection
        ownerId={String(owner.id)}
        ownerTitle={owner.fields.Title}
        ownership={ownership.data ?? []}
        properties={properties.data ?? []}
      />

      {/* Owner-scoped documents — formation docs, EIN, Articles, COE, Cert of Auth, etc. */}
      <EntityDocumentsSection
        ownerIds={Array.from(getAncestorOwnerIds(String(owner.id), ownership.data ?? []))}
        primaryOwnerTitle={owner.fields.Title}
        title="Owner Documents"
        subtitle="This entity's own formation docs PLUS anything tagged to a parent in the ownership chain (e.g. an Assignment of Interest tagged to the holding company shows up here on every subsidiary). Includes the CAHP Entity Documents library."
        uploadOwnerId={String(owner.id)}
        useCahpEntityLibrary
      />

      {/* Deeds where this entity is the grantee */}
      <DeedsSection ownerId={String(owner.id)} ownerTitle={owner.fields.Title} />
      {editingOwnershipId && (
        <EditOwnershipModal
          ownershipId={editingOwnershipId}
          onClose={() => setEditingOwnershipId(null)}
          onSaved={() => ownership.refetch?.()}
        />
      )}
    </div>
  );
}

function KPI({ label, value, sub }: { label: string; value: number; sub?: string }) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className="text-3xl font-bold mt-1 text-teal-700">{value}</div>
      {sub && <div className="text-xs text-gray-500 mt-1">{sub}</div>}
    </div>
  );
}

function maskTaxID(taxId: string | undefined): string {
  if (!taxId) return '';
  if (taxId.length <= 4) return '••••';
  return '••••••' + taxId.slice(-4);
}

// ─────────────────────────────────────────────────────────────────────────────
// Outstanding Items for this owner — spans every property the owner holds
// (direct or through nested LLCs). Has a filter to narrow to items assigned
// to this owner (by name OR contact email) and an Export button that hands
// off to the shared reminder modal so you can copy/paste a ping to them.
// ─────────────────────────────────────────────────────────────────────────────
type OwnerItemsFilter = 'all' | 'assignedToOwner' | 'unassigned';

function OwnerOutstandingItemsSection({
  ownerId,
  ownerTitle,
  ownership,
  properties,
}: {
  ownerId: string;
  ownerTitle: string | undefined;
  ownership: Ownership[];
  properties: Property[];
}) {
  const items = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const contactOwnerLinks = useSharePointList<ContactOwnerLink>(LIST_NAMES.ContactOwnerLinks, { top: 2000 });
  const [filter, setFilter] = useState<OwnerItemsFilter>('all');
  const [exportOpen, setExportOpen] = useState(false);

  // Resolve every property ID this owner has a beneficial interest in
  const ownedPropertyIds = useMemo(
    () => getPropertyIdsForOwner(ownerId, ownership),
    [ownerId, ownership],
  );

  const propertiesById = useMemo(() => {
    const m = new Map<string, Property>();
    for (const p of properties) m.set(String(p.id), p);
    return m;
  }, [properties]);

  const ownedProperties = useMemo(
    () => Array.from(ownedPropertyIds).map((id) => propertiesById.get(id)).filter((p): p is Property => !!p),
    [ownedPropertyIds, propertiesById],
  );

  const isClosed = (s: string | undefined) =>
    s === 'Done' || s === 'Received' || s === 'Not Applicable';

  // Match AssignedTo against any Contact that's linked to this Owner.
  // Source of truth = ContactOwnerLinks junction (many-to-many); also includes
  // the legacy single ContactOwnerLookupId field for back-compat. Compare
  // both the contact's name and email so items assigned either way match.
  const linkedContactKeys = useMemo(() => {
    const linkedContactIds = new Set<string>();
    // Junction: collect contact IDs whose row points at this owner
    for (const link of contactOwnerLinks.data ?? []) {
      if (String(link.fields.OwnerLookupId ?? '') !== String(ownerId)) continue;
      if (link.fields.ContactLookupId) {
        linkedContactIds.add(String(link.fields.ContactLookupId));
      }
    }
    // Legacy single-owner field
    for (const c of contacts.data ?? []) {
      if (String(c.fields.ContactOwnerLookupId ?? '') === String(ownerId)) {
        linkedContactIds.add(String(c.id));
      }
    }

    const keys = new Set<string>();
    for (const c of contacts.data ?? []) {
      if (!linkedContactIds.has(String(c.id))) continue;
      const name = (c.fields.Title ?? '').trim().toLowerCase();
      const email = (c.fields.ContactEmail ?? '').trim().toLowerCase();
      if (name) keys.add(name);
      if (email) keys.add(email);
    }
    // Backward-compat: also match against the Owner entity's own name (legacy items)
    const ownerName = (ownerTitle ?? '').trim().toLowerCase();
    if (ownerName) keys.add(ownerName);
    return keys;
  }, [contacts.data, contactOwnerLinks.data, ownerId, ownerTitle]);

  const isAssignedToThisOwner = (assignedTo: string | undefined) => {
    const a = (assignedTo ?? '').trim().toLowerCase();
    if (!a) return false;
    return linkedContactKeys.has(a);
  };

  // Filter the list against the owned-property set + status + assignee filter
  const filteredItems = useMemo(() => {
    const all = items.data ?? [];
    return all
      .filter((i) => {
        const pid = i.fields.PropertyLookupId ? String(i.fields.PropertyLookupId) : '';
        if (!ownedPropertyIds.has(pid)) return false;
        if (isClosed(i.fields.ItemStatus)) return false;
        if (filter === 'assignedToOwner' && !isAssignedToThisOwner(i.fields.AssignedTo)) return false;
        if (filter === 'unassigned' && (i.fields.AssignedTo ?? '').trim() !== '') return false;
        return true;
      })
      .sort((a, b) => {
        const aOverdue = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() < Date.now() : false;
        const bOverdue = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() < Date.now() : false;
        if (aOverdue !== bOverdue) return aOverdue ? -1 : 1;
        const aDue = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Infinity;
        const bDue = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Infinity;
        return aDue - bDue;
      });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items.data, ownedPropertyIds, filter, linkedContactKeys]);

  // Group filtered items by property for the rendered list
  const itemsByProperty = useMemo(() => {
    const groups = new Map<string, { property: Property | undefined; items: OutstandingItem[] }>();
    for (const item of filteredItems) {
      const pid = item.fields.PropertyLookupId ? String(item.fields.PropertyLookupId) : '';
      if (!groups.has(pid)) {
        groups.set(pid, { property: propertiesById.get(pid), items: [] });
      }
      groups.get(pid)!.items.push(item);
    }
    return Array.from(groups.values()).sort((a, b) =>
      (a.property?.fields.Title ?? '').localeCompare(b.property?.fields.Title ?? ''),
    );
  }, [filteredItems, propertiesById]);

  const assignedCount = useMemo(
    () => (items.data ?? []).filter((i) => {
      const pid = i.fields.PropertyLookupId ? String(i.fields.PropertyLookupId) : '';
      return ownedPropertyIds.has(pid) && !isClosed(i.fields.ItemStatus) && isAssignedToThisOwner(i.fields.AssignedTo);
    }).length,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [items.data, ownedPropertyIds, linkedContactKeys],
  );

  const totalOpenAcrossProperties = useMemo(
    () => (items.data ?? []).filter((i) => {
      const pid = i.fields.PropertyLookupId ? String(i.fields.PropertyLookupId) : '';
      return ownedPropertyIds.has(pid) && !isClosed(i.fields.ItemStatus);
    }).length,
    [items.data, ownedPropertyIds],
  );

  return (
    <div className="bg-white border-l-4 border-amber-500 border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100 bg-amber-50 flex items-start justify-between gap-3 flex-wrap">
        <div>
          <h3 className="text-sm font-semibold text-teal-900">
            Outstanding Items {ownerTitle && <span className="text-gray-500 font-normal">— {ownerTitle}</span>}
          </h3>
          <p className="text-xs text-gray-600 mt-0.5">
            Across <strong>{ownedProperties.length}</strong> propert{ownedProperties.length === 1 ? 'y' : 'ies'} this owner holds
            {' · '}<strong>{totalOpenAcrossProperties}</strong> open total
            {assignedCount > 0 && <> · <strong>{assignedCount}</strong> waiting on this owner</>}
          </p>
        </div>
        <button
          onClick={() => setExportOpen(true)}
          disabled={filteredItems.length === 0}
          className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5 flex-shrink-0"
          title="Export the current filtered list as a copy-pastable reminder"
        >
          <Icon name="file" size={12} />
          Export reminder
        </button>
      </div>

      <div className="px-4 py-2 border-b border-gray-100 bg-gray-50 flex items-center gap-2 flex-wrap">
        <span className="text-[11px] font-semibold text-gray-600 uppercase tracking-wider">Filter:</span>
        {([
          { id: 'all',              label: `All open (${totalOpenAcrossProperties})` },
          { id: 'assignedToOwner',  label: `Waiting on this owner (${assignedCount})` },
          { id: 'unassigned',       label: 'Unassigned' },
        ] as const).map((opt) => (
          <button
            key={opt.id}
            onClick={() => setFilter(opt.id)}
            className={`px-2 py-1 rounded text-[11px] font-medium transition-colors ${
              filter === opt.id
                ? 'bg-teal-700 text-white'
                : 'bg-white text-gray-700 border border-gray-200 hover:bg-gray-100'
            }`}
          >
            {opt.label}
          </button>
        ))}
      </div>

      {ownedProperties.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500 italic">
          This owner doesn't hold any properties yet.
        </div>
      ) : itemsByProperty.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500 italic">
          No items match this filter.
        </div>
      ) : (
        <div className="divide-y divide-gray-100">
          {itemsByProperty.map((group) => (
            <div key={group.property?.id ?? 'unknown'}>
              <div className="px-4 py-2 bg-gray-50 flex items-center justify-between gap-2">
                <Link
                  to={group.property ? `/properties/${group.property.id}` : '#'}
                  className="text-xs font-semibold text-teal-700 hover:text-teal-900"
                >
                  {group.property?.fields.Title ?? '(unknown property)'}
                </Link>
                <span className="text-[11px] text-gray-500 font-mono-data">
                  {group.items.length} open
                </span>
              </div>
              <ul className="divide-y divide-gray-100">
                {group.items.map((item) => {
                  const due = item.fields.DueDate ? new Date(item.fields.DueDate) : null;
                  const isOverdue = !!due && due.getTime() < Date.now();
                  return (
                    <li key={item.id}>
                      <Link
                        to={`/outstanding-items/${item.id}`}
                        className="px-4 py-2 flex items-center gap-3 hover:bg-gray-50 text-xs"
                      >
                        <div className="flex-1 min-w-0">
                          <div className="font-medium text-gray-900 truncate">
                            {isOverdue && <span className="text-error mr-1">⚠</span>}
                            {item.fields.Title}
                          </div>
                          <div className="text-[11px] text-gray-500 mt-0.5">
                            {item.fields.ItemCategory ?? 'Other'}
                            {item.fields.AssignedTo && <span> · {item.fields.AssignedTo}</span>}
                            {item.fields.ItemStatus && <span> · {item.fields.ItemStatus}</span>}
                          </div>
                        </div>
                        <div className="text-right flex-shrink-0">
                          {due ? (
                            <span className={`font-mono-data ${isOverdue ? 'text-error font-bold' : 'text-gray-700'}`}>
                              {formatDateOnly(item.fields.DueDate)}
                            </span>
                          ) : (
                            <span className="text-gray-400">no due date</span>
                          )}
                        </div>
                      </Link>
                    </li>
                  );
                })}
              </ul>
            </div>
          ))}
        </div>
      )}

      {exportOpen && (
        <ExportOutstandingItemsModal
          items={filteredItems}
          propertiesById={propertiesById}
          propertyTitle={undefined /* spans multiple properties — let the modal show all */}
          onClose={() => setExportOpen(false)}
        />
      )}
    </div>
  );
}

// =============================================================================
// CascadePreviewSection — properties affected by member changes (Spec §3.4.2, §5.4)
// =============================================================================
//
// For LLCs and Nonprofits: shows every property where this entity holds direct
// OR indirect interest. Used to warn the user "if you change Stan's % of VanRock,
// these 5 properties are affected."
// =============================================================================

function CascadePreviewSection({
  ownerId,
  ownership,
  owners,
}: {
  ownerId: string;
  ownership: Ownership[];
  owners: Owner[];
}) {
  const navigate = useNavigate();
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const affectedProperties = useMemo(() => {
    if (!properties.data) return [];
    const propsById = new Map(properties.data.map((p) => [String(p.id), p]));

    // Direct: properties where this owner has a direct ownership row
    const directPropertyIds = new Set<string>();
    ownership
      .filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(ownerId) &&
          o.fields.LinkedPropertyLookupId
      )
      .forEach((o) => directPropertyIds.add(String(o.fields.LinkedPropertyLookupId)));

    // Indirect: walk DOWN through subsidiaries (entities this owner invests in)
    // and collect their direct property holdings. Schema: "X owns Y" → row has
    // OwnerLookupId=X, ParentOwnerLookupId=Y. So Y (the subsidiary) is the
    // ParentOwnerLookupId of rows where OwnerLookupId=X.
    const visited = new Set<string>([String(ownerId)]);
    const indirectPropertyIds = new Set<string>();

    function walkDown(currentOwnerId: string) {
      const investments = ownership.filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(currentOwnerId) &&
          o.fields.ParentOwnerLookupId
      );
      for (const rel of investments) {
        const subId = rel.fields.ParentOwnerLookupId;
        if (!subId || visited.has(String(subId))) continue;
        visited.add(String(subId));

        ownership
          .filter(
            (o) =>
              String(o.fields.OwnerLookupId) === String(subId) &&
              o.fields.LinkedPropertyLookupId
          )
          .forEach((o) =>
            indirectPropertyIds.add(String(o.fields.LinkedPropertyLookupId))
          );

        walkDown(String(subId));
      }
    }
    walkDown(String(ownerId));

    // Remove direct from indirect (avoid double-count)
    directPropertyIds.forEach((id) => indirectPropertyIds.delete(id));

    return [
      ...Array.from(directPropertyIds).map((id) => ({
        property: propsById.get(id),
        kind: 'direct' as const,
      })),
      ...Array.from(indirectPropertyIds).map((id) => ({
        property: propsById.get(id),
        kind: 'indirect' as const,
      })),
    ].filter((x) => x.property);
  }, [ownerId, ownership, owners, properties.data]);

  if (affectedProperties.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">Cascade Preview</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            No properties currently affected. When this entity gets ownership stake in a property (directly or via members), affected properties will list here.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100">
        <h3 className="text-sm font-semibold text-gray-700">Cascade Preview</h3>
        <p className="text-xs text-gray-500 mt-0.5">
          Editing this entity's members propagates to{' '}
          <strong>{affectedProperties.length} {affectedProperties.length === 1 ? 'property' : 'properties'}</strong>{' '}
          on the next org chart render. Click any property to navigate.
        </p>
      </div>
      <ul className="divide-y divide-gray-100">
        {affectedProperties.map(({ property, kind }) =>
          property ? (
            <li
              key={property.id}
              onClick={() => navigate(`/properties/${property.id}`)}
              className="px-4 py-2.5 hover:bg-gray-50 cursor-pointer flex items-center justify-between"
            >
              <span className="text-sm font-medium text-gray-900">{property.fields.Title}</span>
              <span
                className={`text-[10px] font-semibold uppercase tracking-wider px-2 py-0.5 rounded ${
                  kind === 'direct'
                    ? 'bg-teal-100 text-teal-800'
                    : 'bg-gray-100 text-gray-600'
                }`}
              >
                {kind === 'direct' ? 'Direct' : 'Indirect'}
              </span>
            </li>
          ) : null
        )}
      </ul>
    </div>
  );
}

// =============================================================================
// OwnerDownstreamTreeSection — informational "everything this entity owns" view.
// Walks the ownership graph DOWN from the current owner to surface every
// subsidiary entity + every property held anywhere in the chain. Reference
// only — DOR filings still use the property-rooted org chart.
// =============================================================================

function OwnerDownstreamTreeSection({
  ownerId,
  ownerTitle,
  ownership,
  owners,
  properties,
}: {
  ownerId: string;
  ownerTitle: string | undefined;
  ownership: Ownership[];
  owners: Owner[];
  properties: Property[];
}) {
  const propertiesById = useMemo(() => {
    const m = new Map<string, Property>();
    for (const p of properties) m.set(String(p.id), p);
    return m;
  }, [properties]);

  const directPropertyIds = useMemo(() => {
    return ownership
      .filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(ownerId) &&
          o.fields.LinkedPropertyLookupId,
      )
      .map((o) => String(o.fields.LinkedPropertyLookupId));
  }, [ownership, ownerId]);

  const tree = useMemo(
    () => getDownstreamTree(ownerId, ownership, owners),
    [ownerId, ownership, owners],
  );

  const totalEntities = useMemo(() => {
    let n = 0;
    function count(nodes: DownstreamNode[]) {
      for (const node of nodes) {
        n++;
        count(node.children);
      }
    }
    count(tree);
    return n;
  }, [tree]);

  const totalProperties = useMemo(() => {
    const seen = new Set<string>(directPropertyIds);
    function walk(nodes: DownstreamNode[]) {
      for (const node of nodes) {
        node.directPropertyIds.forEach((id) => seen.add(id));
        walk(node.children);
      }
    }
    walk(tree);
    return seen.size;
  }, [tree, directPropertyIds]);

  // Empty state — no direct properties AND no subsidiaries. Skip the section
  // entirely so individuals / leaf entities don't get a useless empty card.
  if (directPropertyIds.length === 0 && tree.length === 0) return null;

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100 bg-gray-50">
        <h3 className="text-sm font-semibold text-gray-700">Ownership Tree — Everything Below {ownerTitle}</h3>
        <p className="text-xs text-gray-500 mt-0.5">
          {totalEntities} sub-entit{totalEntities === 1 ? 'y' : 'ies'} · {totalProperties} propert{totalProperties === 1 ? 'y' : 'ies'} held anywhere in the chain.
          <span className="ml-1 italic">Reference view — DOR filings still use each property's own org chart.</span>
        </p>
      </div>
      <div className="p-4 text-sm">
        <DownstreamTreeNode
          ownerName={ownerTitle ?? '(this entity)'}
          ownerType={undefined}
          ownerId={String(ownerId)}
          percentFromParent={undefined}
          directPropertyIds={directPropertyIds}
          children_={tree}
          propertiesById={propertiesById}
          depth={0}
          isRoot
        />
      </div>
    </div>
  );
}

function DownstreamTreeNode({
  ownerName,
  ownerType,
  ownerId,
  percentFromParent,
  directPropertyIds,
  children_,
  propertiesById,
  depth,
  isRoot,
}: {
  ownerName: string;
  ownerType: string | undefined;
  ownerId: string | undefined;
  percentFromParent: number | undefined;
  directPropertyIds: string[];
  children_: DownstreamNode[];
  propertiesById: Map<string, Property>;
  depth: number;
  isRoot?: boolean;
}) {
  const navigate = useNavigate();
  return (
    <div className={depth === 0 ? '' : 'ml-5 pl-3 border-l-2 border-gray-200'}>
      <div className="flex items-center gap-2 py-1">
        <Icon name={isRoot ? 'home' : 'folder'} size={14} className={isRoot ? 'text-teal-700' : 'text-gray-500'} />
        {ownerId ? (
          <button
            onClick={() => navigate(`/owners/${ownerId}`)}
            className={`font-medium text-left ${isRoot ? 'text-teal-900' : 'text-teal-700 hover:text-teal-900 hover:underline'}`}
          >
            {ownerName}
          </button>
        ) : (
          <span className="font-medium text-gray-900">{ownerName}</span>
        )}
        {ownerType && (
          <span className="text-[10px] uppercase tracking-wider text-gray-500 font-semibold">{ownerType}</span>
        )}
        {percentFromParent != null && (
          <span className="text-[11px] text-gray-500 font-mono-data">{percentFromParent}% held from above</span>
        )}
      </div>
      {directPropertyIds.length > 0 && (
        <ul className="ml-6 mb-1">
          {directPropertyIds.map((pid) => {
            const p = propertiesById.get(pid);
            return (
              <li key={pid} className="py-0.5">
                <button
                  onClick={() => navigate(`/properties/${pid}`)}
                  className="text-xs text-gray-700 hover:text-teal-700 hover:underline flex items-center gap-1.5"
                >
                  <Icon name="home" size={11} className="text-gray-400" />
                  {p?.fields.Title ?? `Property #${pid}`}
                  {p?.fields.cahpState && <span className="text-[10px] text-gray-400 font-mono-data">{p.fields.cahpState}</span>}
                </button>
              </li>
            );
          })}
        </ul>
      )}
      {children_.length > 0 && (
        <div>
          {children_.map((child) => (
            <DownstreamTreeNode
              key={`${child.owner.id}-${depth}`}
              ownerName={child.owner.fields.Title ?? '(unnamed)'}
              ownerType={child.owner.fields.OwnerType}
              ownerId={String(child.owner.id)}
              percentFromParent={child.percentFromParent}
              directPropertyIds={child.directPropertyIds}
              children_={child.children}
              propertiesById={propertiesById}
              depth={depth + 1}
            />
          ))}
        </div>
      )}
    </div>
  );
}
