import { Link } from 'react-router-dom';
import { Icon } from '../ui/Icon';
import { useSharePointList, LIST_NAMES, type Submittal } from '../../lib/sharepoint';

export const SC_FILING_FREEZE = new Date('2026-06-25T23:59:59');

export interface FilingFreezeStatus {
  propertyCount: number;
  parcelCount: number;
  daysLeft: number;
  isPastFreeze: boolean;
}

export function useFilingFreezeStatus(): FilingFreezeStatus | null {
  const { data: submittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  if (!submittals) return null;

  const drafts = submittals.filter(
    (s) =>
      s.fields.cahpState === 'SC' &&
      (s.fields.SubmittalStatus === 'Draft' || !s.fields.SubmittalStatus),
  );
  if (drafts.length === 0) return null;

  const properties = new Set<string>();
  const parcels = new Set<string>();
  for (const s of drafts) {
    if (s.fields.PropertyLookupId) properties.add(String(s.fields.PropertyLookupId));
    if (s.fields.TaxMapIDLookupId) parcels.add(String(s.fields.TaxMapIDLookupId));
  }

  const daysLeft = Math.ceil((SC_FILING_FREEZE.getTime() - Date.now()) / 86_400_000);
  return {
    propertyCount: properties.size,
    parcelCount: parcels.size,
    daysLeft,
    isPastFreeze: daysLeft < 0,
  };
}

interface FilingFreezeBannerProps {
  status: FilingFreezeStatus;
}

export function FilingFreezeBanner({ status }: FilingFreezeBannerProps) {
  const { propertyCount, parcelCount, daysLeft, isPastFreeze } = status;

  const palette = isPastFreeze ? 'bg-red-600 text-white' : 'bg-amber-500 text-amber-950';
  const buttonClass = isPastFreeze
    ? 'bg-white text-red-700 hover:bg-red-50'
    : 'bg-amber-950 text-amber-50 hover:bg-amber-900';

  const scopeLabel = parcelCount > 0
    ? `${parcelCount} parcel${parcelCount === 1 ? '' : 's'}${propertyCount > 0 ? ` across ${propertyCount} propert${propertyCount === 1 ? 'y' : 'ies'}` : ''}`
    : `${propertyCount} propert${propertyCount === 1 ? 'y' : 'ies'}`;

  const deadlineLabel = isPastFreeze
    ? 'past the June 25 SC freeze'
    : daysLeft === 0
      ? 'SC freeze starts TODAY'
      : `${daysLeft} day${daysLeft === 1 ? '' : 's'} until June 25 SC freeze`;

  return (
    <div
      className={`fixed top-14 left-0 right-0 z-30 h-10 ${palette} shadow-sm flex items-center px-4 gap-3`}
      role="alert"
    >
      <Icon name="alert" size={16} />
      <div className="flex-1 min-w-0 text-sm font-semibold truncate">
        <span className="font-mono-data">{scopeLabel}</span>{' '}
        <span className="font-normal">still in Draft</span>
        <span className="mx-2 opacity-50">·</span>
        <span>{deadlineLabel}</span>
      </div>
      <Link
        to="/submittals"
        className={`flex-shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded text-xs font-semibold ${buttonClass} transition-colors`}
      >
        Review Drafts
        <Icon name="chevron-right" size={12} />
      </Link>
    </div>
  );
}
