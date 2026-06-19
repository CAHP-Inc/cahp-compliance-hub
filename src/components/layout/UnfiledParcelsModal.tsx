import { useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { Icon } from '../ui/Icon';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Submittal,
  type TaxMapID,
  type Property,
} from '../../lib/sharepoint';
import { CURRENT_FILING_YEAR } from './FilingFreezeBanner';

interface UnfiledParcelsModalProps {
  onClose: () => void;
}

interface ParcelRow {
  parcel: TaxMapID;
  /** Existing Draft submittal for this parcel + current year, if one exists */
  draftSubmittal: Submittal | null;
}

interface PropertyGroup {
  propertyId: string;
  propertyTitle: string;
  rows: ParcelRow[];
}

export function UnfiledParcelsModal({ onClose }: UnfiledParcelsModalProps) {
  const submittalsQ = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const taxMapIDsQ = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const propertiesQ = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const [bulkCreating, setBulkCreating] = useState(false);
  const [bulkProgress, setBulkProgress] = useState<string | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);
  const [draftFilter, setDraftFilter] = useState<'all' | 'draft' | 'none'>('all');

  const groupedByProperty = useMemo<PropertyGroup[]>(() => {
    if (!submittalsQ.data || !taxMapIDsQ.data || !propertiesQ.data) return [];

    const propertiesById = new Map<string, Property>(
      propertiesQ.data.map((p) => [String(p.id), p]),
    );

    const scPropertyIds = new Set(
      propertiesQ.data.filter((p) => p.fields.cahpState === 'SC').map((p) => String(p.id)),
    );

    // "Filed" = any submittal with non-Draft status referencing this parcel.
    const filedParcelIds = new Set<string>();
    for (const s of submittalsQ.data) {
      const tmid = s.fields.TaxMapIDLookupId ? String(s.fields.TaxMapIDLookupId) : '';
      if (!tmid) continue;
      const status = s.fields.SubmittalStatus;
      if (status && status !== 'Draft') filedParcelIds.add(tmid);
    }

    // Existing Draft submittals for the current filing year, keyed by parcel id.
    const draftByParcel = new Map<string, Submittal>();
    for (const s of submittalsQ.data) {
      const tmid = s.fields.TaxMapIDLookupId ? String(s.fields.TaxMapIDLookupId) : '';
      if (!tmid) continue;
      if (s.fields.cahpTaxYear !== CURRENT_FILING_YEAR) continue;
      const status = s.fields.SubmittalStatus;
      if (status && status !== 'Draft') continue;
      // Prefer existing draftSubmittal already in map (first-seen wins)
      if (!draftByParcel.has(tmid)) draftByParcel.set(tmid, s);
    }

    const groups = new Map<string, PropertyGroup>();
    for (const t of taxMapIDsQ.data) {
      const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
      if (!pid || !scPropertyIds.has(pid)) continue;
      if (filedParcelIds.has(String(t.id))) continue;

      const propertyTitle = propertiesById.get(pid)?.fields.Title ?? `Property #${pid}`;
      if (!groups.has(pid)) {
        groups.set(pid, { propertyId: pid, propertyTitle, rows: [] });
      }
      groups.get(pid)!.rows.push({
        parcel: t,
        draftSubmittal: draftByParcel.get(String(t.id)) ?? null,
      });
    }

    // Sort rows by parcel title within each group
    for (const g of groups.values()) {
      g.rows.sort((a, b) => (a.parcel.fields.Title ?? '').localeCompare(b.parcel.fields.Title ?? ''));
    }

    return Array.from(groups.values()).sort((a, b) => a.propertyTitle.localeCompare(b.propertyTitle));
  }, [submittalsQ.data, taxMapIDsQ.data, propertiesQ.data]);

  // Filter rows by Draft state: all / has a Draft started / no submittal yet.
  const visibleGroups = useMemo<PropertyGroup[]>(() => {
    if (draftFilter === 'all') return groupedByProperty;
    return groupedByProperty
      .map((g) => ({
        ...g,
        rows: g.rows.filter((r) => (draftFilter === 'draft' ? r.draftSubmittal : !r.draftSubmittal)),
      }))
      .filter((g) => g.rows.length > 0);
  }, [groupedByProperty, draftFilter]);

  const totalParcels = groupedByProperty.reduce((sum, g) => sum + g.rows.length, 0);
  const draftStartedCount = groupedByProperty.reduce(
    (sum, g) => sum + g.rows.filter((r) => r.draftSubmittal).length,
    0,
  );
  const needsDraftRows = useMemo(
    () => groupedByProperty.flatMap((g) => g.rows.filter((r) => !r.draftSubmittal).map((r) => ({ group: g, row: r }))),
    [groupedByProperty],
  );

  const handleBulkCreate = async () => {
    if (needsDraftRows.length === 0) return;
    if (!window.confirm(`Create ${needsDraftRows.length} Draft submittals for tax year ${CURRENT_FILING_YEAR}? This cannot be batch-undone.`)) {
      return;
    }
    setBulkCreating(true);
    setBulkError(null);
    let created = 0;
    try {
      for (const { group, row } of needsDraftRows) {
        setBulkProgress(`Creating ${created + 1} of ${needsDraftRows.length}…`);
        await createListItem(LIST_NAMES.Submittals, {
          Title: `${group.propertyTitle} — ${row.parcel.fields.Title} — Initial ${CURRENT_FILING_YEAR}`,
          PropertyLookupId: Number(group.propertyId),
          TaxMapIDLookupId: Number(row.parcel.id),
          cahpTaxYear: CURRENT_FILING_YEAR,
          cahpState: 'SC',
          SubmittalStatus: 'Draft',
          FilingType: 'Initial',
        });
        created++;
      }
      setBulkProgress(`Done — created ${created} Draft${created === 1 ? '' : 's'}.`);
      submittalsQ.refetch?.();
      setTimeout(() => {
        setBulkCreating(false);
        setBulkProgress(null);
      }, 1200);
    } catch (err) {
      setBulkError(`Created ${created} of ${needsDraftRows.length} before failing: ${err instanceof Error ? err.message : String(err)}`);
      setBulkCreating(false);
    }
  };

  const loading = !submittalsQ.data || !taxMapIDsQ.data || !propertiesQ.data;

  return (
    <div
      className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-lg shadow-xl max-w-3xl w-full max-h-[80vh] flex flex-col my-8"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="px-5 py-4 border-b border-gray-200 flex items-center justify-between">
          <div>
            <h3 className="text-lg font-bold text-teal-700">Unfiled SC Parcels</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {loading
                ? 'Loading…'
                : `${totalParcels} parcel${totalParcels === 1 ? '' : 's'} across ${groupedByProperty.length} propert${groupedByProperty.length === 1 ? 'y' : 'ies'} not yet filed for ${CURRENT_FILING_YEAR}.`}
            </p>
          </div>
          <button
            onClick={onClose}
            className="text-gray-400 hover:text-gray-700 w-8 h-8 flex items-center justify-center rounded hover:bg-gray-100"
            aria-label="Close"
          >
            <Icon name="x" size={18} />
          </button>
        </div>

        {!loading && needsDraftRows.length > 0 && (
          <div className="px-5 py-3 bg-amber-50 border-b border-amber-200 flex items-center justify-between gap-3">
            <div className="text-xs text-amber-900">
              <strong>{needsDraftRows.length}</strong> parcel{needsDraftRows.length === 1 ? '' : 's'} have no Draft started yet for {CURRENT_FILING_YEAR}.
            </div>
            <button
              onClick={handleBulkCreate}
              disabled={bulkCreating}
              className="flex-shrink-0 inline-flex items-center gap-1.5 px-3 py-1.5 rounded text-xs font-semibold bg-amber-700 text-white hover:bg-amber-800 disabled:bg-gray-400"
            >
              {bulkCreating ? (
                <>
                  <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
                  Creating…
                </>
              ) : (
                <>Create {needsDraftRows.length} Draft{needsDraftRows.length === 1 ? '' : 's'}</>
              )}
            </button>
          </div>
        )}

        {bulkProgress && (
          <div className="px-5 py-2 bg-blue-50 border-b border-blue-200 text-xs text-blue-900 flex items-center gap-2">
            {bulkCreating && (
              <div className="w-3 h-3 rounded-full border-2 border-blue-500 border-r-transparent animate-spin" />
            )}
            {bulkProgress}
          </div>
        )}

        {bulkError && (
          <div className="px-5 py-2 bg-red-50 border-b border-red-200 text-xs text-red-900 flex items-start gap-2">
            <Icon name="alert" size={14} className="flex-shrink-0 mt-0.5" />
            {bulkError}
          </div>
        )}

        {!loading && totalParcels > 0 && (
          <div className="px-5 py-2 border-b border-gray-200 flex items-center gap-1 text-xs">
            <span className="text-gray-500 mr-1">Show:</span>
            {([
              ['all', `All (${totalParcels})`],
              ['draft', `Draft started (${draftStartedCount})`],
              ['none', `No submittal (${totalParcels - draftStartedCount})`],
            ] as const).map(([key, label]) => (
              <button
                key={key}
                onClick={() => setDraftFilter(key)}
                className={`px-2 py-1 rounded border ${draftFilter === key ? 'bg-teal-700 text-white border-teal-700' : 'border-gray-300 text-gray-600 hover:bg-gray-50'}`}
              >
                {label}
              </button>
            ))}
          </div>
        )}

        <div className="flex-1 overflow-y-auto">
          {loading ? (
            <div className="p-8 text-center text-sm text-gray-500">Loading…</div>
          ) : visibleGroups.length === 0 ? (
            <div className="p-8 text-center text-sm text-gray-500">
              {groupedByProperty.length === 0
                ? 'All SC parcels have been filed. Nothing left to do.'
                : 'No parcels match this filter.'}
            </div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {visibleGroups.map((group) => (
                <li key={group.propertyId} className="px-5 py-3">
                  <div className="flex items-center justify-between mb-2">
                    <Link
                      to={`/properties/${group.propertyId}`}
                      onClick={onClose}
                      className="text-sm font-semibold text-gray-900 hover:text-teal-700"
                    >
                      {group.propertyTitle}
                    </Link>
                    <span className="text-xs text-gray-500 font-mono-data">
                      {group.rows.length} parcel{group.rows.length === 1 ? '' : 's'}
                    </span>
                  </div>
                  <ul className="space-y-1 pl-3 border-l-2 border-amber-300">
                    {group.rows.map(({ parcel, draftSubmittal }) => {
                      const target = draftSubmittal
                        ? `/submittals/${draftSubmittal.id}`
                        : `/properties/${group.propertyId}`;
                      return (
                        <li key={parcel.id}>
                          <Link
                            to={target}
                            onClick={onClose}
                            className="flex items-center gap-3 py-1.5 px-2 rounded hover:bg-amber-50 transition-colors text-sm"
                          >
                            <span className="font-mono-data text-gray-900 flex-shrink-0">
                              {parcel.fields.Title}
                            </span>
                            {parcel.fields.ParcelAddress && (
                              <span className="text-xs text-gray-500 truncate">
                                {parcel.fields.ParcelAddress}
                              </span>
                            )}
                            <span className="ml-auto inline-flex items-center gap-1.5 text-xs flex-shrink-0">
                              {draftSubmittal ? (
                                <span className="px-1.5 py-0.5 rounded bg-gray-100 text-gray-700 text-[10px] font-semibold">
                                  Draft started
                                </span>
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-red-100 text-red-800 text-[10px] font-semibold">
                                  No submittal
                                </span>
                              )}
                              <Icon name="chevron-right" size={12} className="text-gray-400" />
                            </span>
                          </Link>
                        </li>
                      );
                    })}
                  </ul>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="px-5 py-3 border-t border-gray-200 flex items-center justify-between">
          <Link
            to="/submittals"
            onClick={onClose}
            className="text-xs text-teal-700 hover:text-teal-900 font-medium"
          >
            Open Submittals page →
          </Link>
          <button
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-md"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
