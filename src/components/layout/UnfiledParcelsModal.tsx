import { useMemo } from 'react';
import { Link } from 'react-router-dom';
import { Icon } from '../ui/Icon';
import {
  useSharePointList,
  LIST_NAMES,
  type Submittal,
  type TaxMapID,
  type Property,
} from '../../lib/sharepoint';

interface UnfiledParcelsModalProps {
  onClose: () => void;
}

export function UnfiledParcelsModal({ onClose }: UnfiledParcelsModalProps) {
  const { data: submittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const { data: taxMapIDs } = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
  const { data: properties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });

  const propertiesById = useMemo(() => {
    const m = new Map<string, Property>();
    (properties ?? []).forEach((p) => m.set(String(p.id), p));
    return m;
  }, [properties]);

  const taxMapIdsById = useMemo(() => {
    const m = new Map<string, TaxMapID>();
    (taxMapIDs ?? []).forEach((t) => m.set(String(t.id), t));
    return m;
  }, [taxMapIDs]);

  // Group draft SC submittals by property → list of parcels under each.
  const groupedByProperty = useMemo(() => {
    if (!submittals) return [];
    const drafts = submittals.filter(
      (s) =>
        s.fields.cahpState === 'SC' &&
        (s.fields.SubmittalStatus === 'Draft' || !s.fields.SubmittalStatus),
    );
    const groups = new Map<string, {
      propertyId: string;
      propertyTitle: string;
      rows: { submittal: Submittal; parcel: TaxMapID | null }[];
    }>();
    for (const s of drafts) {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '__unlinked__';
      const propertyTitle = pid === '__unlinked__'
        ? '(no property linked)'
        : propertiesById.get(pid)?.fields.Title ?? `Property #${pid}`;
      if (!groups.has(pid)) {
        groups.set(pid, { propertyId: pid, propertyTitle, rows: [] });
      }
      const parcel = s.fields.TaxMapIDLookupId
        ? taxMapIdsById.get(String(s.fields.TaxMapIDLookupId)) ?? null
        : null;
      groups.get(pid)!.rows.push({ submittal: s, parcel });
    }
    return Array.from(groups.values()).sort((a, b) => a.propertyTitle.localeCompare(b.propertyTitle));
  }, [submittals, propertiesById, taxMapIdsById]);

  const totalParcels = groupedByProperty.reduce((sum, g) => sum + g.rows.length, 0);

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
              {totalParcels} parcel{totalParcels === 1 ? '' : 's'} across {groupedByProperty.length} propert{groupedByProperty.length === 1 ? 'y' : 'ies'} still in Draft for the SC filing window.
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

        <div className="flex-1 overflow-y-auto">
          {!submittals ? (
            <div className="p-8 text-center text-sm text-gray-500">Loading…</div>
          ) : groupedByProperty.length === 0 ? (
            <div className="p-8 text-center text-sm text-gray-500">
              All SC parcels have been filed. Nothing left in Draft.
            </div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {groupedByProperty.map((group) => (
                <li key={group.propertyId} className="px-5 py-3">
                  <div className="flex items-center justify-between mb-2">
                    <div className="text-sm font-semibold text-gray-900">{group.propertyTitle}</div>
                    <span className="text-xs text-gray-500 font-mono-data">
                      {group.rows.length} parcel{group.rows.length === 1 ? '' : 's'}
                    </span>
                  </div>
                  <ul className="space-y-1 pl-3 border-l-2 border-amber-300">
                    {group.rows.map(({ submittal, parcel }) => (
                      <li key={submittal.id}>
                        <Link
                          to={`/submittals/${submittal.id}`}
                          onClick={onClose}
                          className="flex items-center gap-3 py-1.5 px-2 rounded hover:bg-amber-50 transition-colors text-sm"
                        >
                          <span className="font-mono-data text-gray-900 flex-shrink-0">
                            {parcel?.fields.Title ?? submittal.fields.Title ?? '(no parcel ID)'}
                          </span>
                          {parcel?.fields.ParcelAddress && (
                            <span className="text-xs text-gray-500 truncate">
                              {parcel.fields.ParcelAddress}
                            </span>
                          )}
                          <span className="ml-auto inline-flex items-center gap-1.5 text-xs flex-shrink-0">
                            <span className="px-1.5 py-0.5 rounded bg-gray-100 text-gray-700 text-[10px] font-semibold">
                              Draft
                            </span>
                            {submittal.fields.cahpTaxYear && (
                              <span className="text-gray-500 font-mono-data">
                                {submittal.fields.cahpTaxYear}
                              </span>
                            )}
                            <Icon name="chevron-right" size={12} className="text-gray-400" />
                          </span>
                        </Link>
                      </li>
                    ))}
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
