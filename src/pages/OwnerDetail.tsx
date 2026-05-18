import { useState, useEffect, useMemo } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
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
  type Owner,
  type OwnerFields,
  type Ownership,
  type Property,
  type OwnerType,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { EntityDocumentsSection } from '../components/EntityDocumentsSection';
import { EditOwnershipModal } from '../components/EditOwnershipModal';
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
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">Property Holdings</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            {directProperties.length} direct property {directProperties.length === 1 ? 'holding' : 'holdings'}
          </p>
        </div>
        {directProperties.length === 0 ? (
          <div className="px-4 py-8 text-center text-sm text-gray-500">
            This owner doesn't directly hold any properties.
            {stats && stats.propertyCounts.indirect > 0 && (
              <> Has indirect interest in <strong>{stats.propertyCounts.indirect}</strong> propert{stats.propertyCounts.indirect === 1 ? 'y' : 'ies'} via owned entities.</>
            )}
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

      {/* Owner-scoped documents — formation docs, EIN, Articles, COE, Cert of Auth, etc. */}
      <EntityDocumentsSection
        ownerIds={[String(owner.id)]}
        primaryOwnerTitle={owner.fields.Title}
        title="Owner Documents"
        subtitle="Entity formation docs (EIN, Articles, COE, Cert of Authorization, OA) — uploaded once, surfaced on every property this owner holds."
        uploadOwnerId={String(owner.id)}
      />

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

    // Indirect: walk DOWN to find properties held by entities this owner has stake in
    const visited = new Set<string>([String(ownerId)]);
    const indirectPropertyIds = new Set<string>();

    function walkDown(currentOwnerId: string) {
      const childRels = ownership.filter(
        (o) => String(o.fields.ParentOwnerLookupId) === String(currentOwnerId)
      );
      for (const rel of childRels) {
        const childOwnerId = rel.fields.OwnerLookupId;
        if (!childOwnerId || visited.has(String(childOwnerId))) continue;
        visited.add(String(childOwnerId));

        ownership
          .filter(
            (o) =>
              String(o.fields.OwnerLookupId) === String(childOwnerId) &&
              o.fields.LinkedPropertyLookupId
          )
          .forEach((o) =>
            indirectPropertyIds.add(String(o.fields.LinkedPropertyLookupId))
          );

        walkDown(String(childOwnerId));
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
