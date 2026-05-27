import { useState } from 'react';
import { Icon } from '../ui/Icon';
import {
  useSharePointList,
  LIST_NAMES,
  type Submittal,
  type TaxMapID,
  type Property,
} from '../../lib/sharepoint';
import { UnfiledParcelsModal } from './UnfiledParcelsModal';

export const SC_FILING_FREEZE = new Date('2026-06-25T23:59:59-04:00');
export const CURRENT_FILING_YEAR = '2026';

export interface FilingFreezeStatus {
  propertyCount: number;
  parcelCount: number;
  daysLeft: number;
  isPastFreeze: boolean;
}

/**
 * "Unfiled" mirrors the Properties page rule: a parcel is filed if any
 * submittal references it with a status other than Draft / blank. Unfiled =
 * total SC parcels minus filed SC parcels. This counts parcels with NO
 * submittal at all in addition to parcels stuck in Draft.
 */
export function useFilingFreezeStatus(): FilingFreezeStatus | null {
  const { data: submittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const { data: taxMapIDs } = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const { data: properties } = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  if (!submittals || !taxMapIDs || !properties) return null;

  const scPropertyIds = new Set(
    properties.filter((p) => p.fields.cahpState === 'SC').map((p) => String(p.id)),
  );

  const filedParcelIds = new Set<string>();
  for (const s of submittals) {
    const tmid = s.fields.TaxMapIDLookupId ? String(s.fields.TaxMapIDLookupId) : '';
    if (!tmid) continue;
    const status = s.fields.SubmittalStatus;
    if (status && status !== 'Draft') filedParcelIds.add(tmid);
  }

  const unfiledProperties = new Set<string>();
  let unfiledParcelCount = 0;
  for (const t of taxMapIDs) {
    const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
    if (!pid || !scPropertyIds.has(pid)) continue;
    if (filedParcelIds.has(String(t.id))) continue;
    unfiledParcelCount++;
    unfiledProperties.add(pid);
  }

  if (unfiledParcelCount === 0) return null;

  const daysLeft = Math.ceil((SC_FILING_FREEZE.getTime() - Date.now()) / 86_400_000);
  return {
    propertyCount: unfiledProperties.size,
    parcelCount: unfiledParcelCount,
    daysLeft,
    isPastFreeze: daysLeft < 0,
  };
}

interface FilingFreezeBannerProps {
  status: FilingFreezeStatus;
}

export function FilingFreezeBanner({ status }: FilingFreezeBannerProps) {
  const [modalOpen, setModalOpen] = useState(false);
  const { propertyCount, parcelCount, daysLeft, isPastFreeze } = status;

  const palette = isPastFreeze ? 'bg-red-600 text-white' : 'bg-amber-500 text-amber-950';
  const buttonClass = isPastFreeze
    ? 'bg-white text-red-700 hover:bg-red-50'
    : 'bg-amber-950 text-amber-50 hover:bg-amber-900';

  const deadlineLabel = isPastFreeze
    ? 'past the June 25 SC freeze'
    : daysLeft === 0
      ? 'SC freeze starts TODAY'
      : `${daysLeft} day${daysLeft === 1 ? '' : 's'} until June 25 SC freeze`;

  return (
    <>
      <div
        className={`fixed top-14 left-0 right-0 z-30 h-10 ${palette} shadow-sm flex items-center px-4 gap-3`}
        role="alert"
      >
        <Icon name="alert" size={16} />
        <div className="flex-1 min-w-0 text-sm font-semibold truncate">
          <span className="font-mono-data">{propertyCount}</span>{' '}
          <span className="font-normal">propert{propertyCount === 1 ? 'y' : 'ies'} still in Draft</span>
          <span className="mx-1.5 opacity-50">-</span>
          <span className="font-mono-data">{parcelCount}</span>{' '}
          <span className="font-normal">parcel{parcelCount === 1 ? '' : 's'} unfiled</span>
          <span className="mx-2 opacity-50">·</span>
          <span>{deadlineLabel}</span>
        </div>
        <button
          onClick={() => setModalOpen(true)}
          className={`flex-shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded text-xs font-semibold ${buttonClass} transition-colors`}
        >
          View Parcels
          <Icon name="chevron-right" size={12} />
        </button>
      </div>
      {modalOpen && <UnfiledParcelsModal onClose={() => setModalOpen(false)} />}
    </>
  );
}
