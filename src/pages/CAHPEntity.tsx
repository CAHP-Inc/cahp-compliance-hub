import { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Owner,
  type Ownership,
  type Property,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { EntityDocumentsSection } from '../components/EntityDocumentsSection';

/**
 * CAHP Entity module — Spec §3.5.
 *
 * A specialized view of the CAHP SC LLC (and future CAHP NC) nonprofit entities.
 * The data source is the standard Owners list — this page just filters and
 * organizes that data in a CAHP-specific way.
 *
 * Entity identification: any Owner with Title containing "CAHP" and OwnerType
 * Nonprofit or LLC. In Brandy's data: "Carolina Affordable Housing Project"
 * (the 501(c)(3) parent) and "CAHP SC LLC" (the disregarded LLC).
 */
export function CAHPEntity() {
  const navigate = useNavigate();
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const loading = owners.loading || ownership.loading || properties.loading;
  const error = owners.error || ownership.error || properties.error;

  // Identify CAHP entities — anything with "CAHP" in the title or
  // explicitly named "Carolina Affordable Housing Project"
  const cahpEntities = useMemo(() => {
    if (!owners.data) return [];
    return owners.data.filter((o) => {
      const title = (o.fields.Title ?? '').toLowerCase();
      return (
        title.includes('cahp') ||
        title.includes('carolina affordable housing project')
      );
    });
  }, [owners.data]);

  const [activeEntityId, setActiveEntityId] = useState<string | null>(null);

  // Default to first entity once data loads
  const effectiveActiveId =
    activeEntityId ?? (cahpEntities.length > 0 ? cahpEntities[0].id : null);

  const activeEntity = useMemo(() => {
    if (!effectiveActiveId) return null;
    return cahpEntities.find((e) => String(e.id) === String(effectiveActiveId)) ?? null;
  }, [cahpEntities, effectiveActiveId]);

  if (loading) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">CAHP Entity</h1>
        <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
          <div className="inline-flex items-center gap-3 text-gray-500">
            <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            <span className="text-sm">Loading CAHP entities…</span>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">CAHP Entity</h1>
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2">Failed to load CAHP entity data</div>
          <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
        </div>
      </div>
    );
  }

  if (cahpEntities.length === 0) {
    return (
      <div>
        <h1 className="text-2xl font-bold text-teal-700 mb-6">CAHP Entity</h1>
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-8 text-center">
          <p className="text-base font-semibold text-blue-900 mb-1">No CAHP entities found</p>
          <p className="text-sm text-blue-800 mb-4">
            This view auto-detects entities by name. To populate it, create Owners named{' '}
            <strong>Carolina Affordable Housing Project</strong> (Nonprofit) and{' '}
            <strong>CAHP SC LLC</strong> (LLC) in the Owners module.
          </p>
          <button
            onClick={() => navigate('/owners/new')}
            className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium inline-flex items-center gap-2 transition-colors"
          >
            <Icon name="plus" size={16} />
            Create CAHP Owner
          </button>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">CAHP Entity</h1>
        <p className="text-sm text-gray-500 mt-1">
          Specialized view of CAHP-named nonprofit and LLC entities. Formation docs, OAs, EINs, and determination letters live in the <strong>Documents</strong> section below.
        </p>
      </div>

      {/* Entity tabs (one per CAHP entity, currently 1-2) */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
        <div className="flex border-b border-gray-200">
          {cahpEntities.map((entity) => {
            const isActive = String(entity.id) === String(effectiveActiveId);
            return (
              <button
                key={entity.id}
                onClick={() => setActiveEntityId(entity.id)}
                className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
                  isActive
                    ? 'border-teal-700 text-teal-700 bg-teal-50'
                    : 'border-transparent text-gray-600 hover:text-gray-900 hover:bg-gray-50'
                }`}
              >
                {entity.fields.Title}
                {entity.fields.OwnerType && (
                  <span className="ml-2 text-[10px] uppercase tracking-wider text-gray-500">
                    {entity.fields.OwnerType}
                  </span>
                )}
              </button>
            );
          })}
        </div>

        {activeEntity && (
          <CAHPEntityTab
            entity={activeEntity}
            ownership={ownership.data ?? []}
            owners={owners.data ?? []}
            properties={properties.data ?? []}
          />
        )}
      </div>
    </div>
  );
}

function CAHPEntityTab({
  entity,
  ownership,
  owners,
  properties,
}: {
  entity: Owner;
  ownership: Ownership[];
  owners: Owner[];
  properties: Property[];
}) {
  const navigate = useNavigate();

  // Property memberships — every property where this CAHP entity is a direct member
  const propertyMemberships = useMemo(() => {
    const propsById = new Map(properties.map((p) => [String(p.id), p]));
    return ownership
      .filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(entity.id) &&
          o.fields.LinkedPropertyLookupId
      )
      .map((o) => ({
        relationship: o,
        property: propsById.get(String(o.fields.LinkedPropertyLookupId)) ?? null,
      }))
      .filter((x) => x.property);
  }, [entity.id, ownership, properties]);

  // Members of this entity — owners that have a ParentOwnerLookupId pointing here
  const members = useMemo(() => {
    const ownersById = new Map(owners.map((o) => [String(o.id), o]));
    return ownership
      .filter((o) => String(o.fields.ParentOwnerLookupId) === String(entity.id))
      .map((o) => ({
        relationship: o,
        owner: o.fields.OwnerLookupId
          ? ownersById.get(String(o.fields.OwnerLookupId)) ?? null
          : null,
      }))
      .filter((x) => x.owner);
  }, [entity.id, ownership, owners]);

  // Entities this owner is a member of (typically just the parent, e.g., CAHP SC LLC → CAHP, Inc.)
  const memberOf = useMemo(() => {
    const ownersById = new Map(owners.map((o) => [String(o.id), o]));
    return ownership
      .filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(entity.id) &&
          o.fields.ParentOwnerLookupId
      )
      .map((o) => ({
        relationship: o,
        parent: o.fields.ParentOwnerLookupId
          ? ownersById.get(String(o.fields.ParentOwnerLookupId)) ?? null
          : null,
      }))
      .filter((x) => x.parent);
  }, [entity.id, ownership, owners]);

  return (
    <div className="p-5">
      {/* Identity */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <KPI label="Entity Name" value={entity.fields.Title ?? '—'} isText />
        <KPI label="EIN" value={maskTax(entity.fields.TaxID)} isText mono />
        <KPI label="State" value={entity.fields.OwnerState ?? '—'} isText mono />
        <KPI label="Members" value={String(members.length)} isText />
      </div>

      {/* Member-of (e.g., CAHP SC LLC is member of CAHP, Inc.) */}
      {memberOf.length > 0 && (
        <div className="mb-6 bg-gold-50 border border-gold-200 rounded-lg p-4">
          <div className="flex items-start gap-3">
            <Icon name="alert" size={16} className="text-gold-700 flex-shrink-0 mt-0.5" />
            <div className="flex-1 text-sm">
              <strong>{entity.fields.Title}</strong> is a member of:{' '}
              {memberOf.map(({ relationship, parent }, idx) => (
                <span key={relationship.id}>
                  {idx > 0 && ', '}
                  <button
                    onClick={() => navigate(`/owners/${parent?.id}`)}
                    className="text-teal-700 hover:text-teal-900 font-medium underline"
                  >
                    {parent?.fields.Title}
                  </button>
                  {' '}({relationship.fields.RelationshipType ?? 'Member'}, {relationship.fields.OwnershipPercent ?? 0}%)
                </span>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Document libraries moved to EntityDocumentsSection (PR-14-hardening) — rendered elsewhere on this page */}

      {/* Members section */}
      {members.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
          <div className="px-4 py-3 border-b border-gray-100">
            <h3 className="text-sm font-semibold text-gray-700">Members of {entity.fields.Title}</h3>
            <p className="text-xs text-gray-500 mt-0.5">{members.length} {members.length === 1 ? 'member' : 'members'}</p>
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
              {members.map(({ relationship, owner }) => (
                <tr
                  key={relationship.id}
                  onClick={() => owner && navigate(`/owners/${owner.id}`)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">{owner?.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{relationship.fields.RelationshipType ?? '—'}</td>
                  <td className="px-4 py-3 text-right font-mono-data">
                    {relationship.fields.OwnershipPercent != null ? `${relationship.fields.OwnershipPercent}%` : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Property Memberships */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden mb-6">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">Property Memberships</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            {propertyMemberships.length === 0
              ? `No properties currently have ${entity.fields.Title} as a member.`
              : `${propertyMemberships.length} ${propertyMemberships.length === 1 ? 'property where' : 'properties where'} ${entity.fields.Title} holds direct interest.`}
          </p>
        </div>
        {propertyMemberships.length > 0 && (
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
              {propertyMemberships.map(({ relationship, property }) => (
                <tr
                  key={relationship.id}
                  onClick={() => property && navigate(`/properties/${property.id}`)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">{property?.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">
                    <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-teal-100 text-teal-800">
                      {relationship.fields.RelationshipType ?? '—'}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-right font-mono-data">
                    {relationship.fields.OwnershipPercent != null ? `${relationship.fields.OwnershipPercent}%` : '—'}
                  </td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                    {relationship.fields.EffectiveDate
                      ? new Date(relationship.fields.EffectiveDate).toLocaleDateString()
                      : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Documents — read from the dedicated CAHP Entity Documents library */}
      <EntityDocumentsSection
        ownerIds={[String(entity.id)]}
        primaryOwnerTitle={entity.fields.Title}
        title="Entity Documents"
        subtitle="OAs, formation docs, EIN, determination letters. Stored in the CAHP Entity Documents library and surfaced as reference material on every property filing."
        uploadOwnerId={String(entity.id)}
        useCahpEntityLibrary
        strictEntityFilter
      />
    </div>
  );
}

function KPI({
  label,
  value,
  isText,
  mono,
}: {
  label: string;
  value: string;
  isText?: boolean;
  mono?: boolean;
}) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`mt-1 ${isText ? 'text-base font-semibold text-gray-900' : 'text-3xl font-bold text-teal-700'} ${mono ? 'font-mono-data' : ''}`}>
        {value}
      </div>
    </div>
  );
}

function maskTax(taxId: string | undefined): string {
  if (!taxId) return '—';
  if (taxId.length <= 4) return '••••';
  return '••••••' + taxId.slice(-4);
}
