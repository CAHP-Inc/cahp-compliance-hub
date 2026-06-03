import { Icon } from './ui/Icon';
import type { FilingPace } from '../lib/filing-pace';

/**
 * Filing-pace card for My Day. Shows filed/total, a progress bar with a
 * "today's target" marker, an ahead/behind badge, and the per-day rate needed
 * to finish by the target date. All values are live — they re-derive as the
 * total and filed counts change.
 */
export function FilingPaceCard({ pace }: { pace: FilingPace }) {
  const {
    filed,
    total,
    remaining,
    filedPct,
    targetPct,
    delta,
    workingDaysRemaining,
    perDayNeeded,
    targetDateLabel,
    status,
  } = pace;

  const behind = status === 'behind';
  const badge =
    status === 'behind'
      ? { text: `${Math.abs(delta)} behind pace`, cls: 'bg-red-100 text-red-800' }
      : status === 'ahead'
        ? { text: `${delta} ahead of pace`, cls: 'bg-green-100 text-green-800' }
        : status === 'on-pace'
          ? { text: 'On pace', cls: 'bg-teal-100 text-teal-800' }
          : status === 'complete'
            ? { text: 'All filed', cls: 'bg-green-100 text-green-800' }
            : status === 'past-target'
              ? { text: 'Past target date', cls: 'bg-red-100 text-red-800' }
              : { text: 'Not started', cls: 'bg-gray-100 text-gray-700' };

  const barColor = behind ? 'bg-red-500' : status === 'complete' ? 'bg-green-500' : 'bg-teal-600';
  const accent = behind ? 'border-red-500' : 'border-teal-600';

  return (
    <div className={`bg-white border border-gray-200 border-l-4 ${accent} rounded-lg shadow-card mb-6`}>
      <div className="px-5 py-4">
        <div className="flex items-start justify-between gap-3 mb-3">
          <div>
            <h2 className="text-base font-semibold text-teal-700 flex items-center gap-2">
              <Icon name="alert" size={16} />
              Filing Pace
            </h2>
            <p className="text-xs text-gray-500 mt-0.5">
              Goal: all SC parcels filed by {targetDateLabel}
            </p>
          </div>
          <span className={`flex-shrink-0 px-2.5 py-1 rounded-full text-xs font-semibold ${badge.cls}`}>
            {badge.text}
          </span>
        </div>

        <div className="flex items-baseline gap-2 mb-2">
          <span className="text-3xl font-bold text-teal-700 font-mono-data">{filed}</span>
          <span className="text-lg text-gray-400 font-mono-data">/ {total}</span>
          <span className="text-sm text-gray-500">parcels filed ({Math.round(filedPct)}%)</span>
        </div>

        {/* Progress bar with a "today's target" marker */}
        <div className="relative h-3 bg-gray-100 rounded-full overflow-visible mb-1">
          <div
            className={`h-3 rounded-full ${barColor} transition-all`}
            style={{ width: `${Math.min(100, filedPct)}%` }}
          />
          {targetPct > 0 && targetPct < 100 && (
            <div
              className="absolute top-[-3px] h-[18px] w-0.5 bg-gray-700"
              style={{ left: `${targetPct}%` }}
              title={`Today's on-pace target: ${Math.round(targetPct)}%`}
            />
          )}
        </div>
        {targetPct > 0 && targetPct < 100 && (
          <div className="relative h-4 mb-2">
            <span
              className="absolute text-[10px] text-gray-500 -translate-x-1/2 whitespace-nowrap"
              style={{ left: `${targetPct}%` }}
            >
              today’s target
            </span>
          </div>
        )}

        {remaining > 0 ? (
          <p className="text-sm text-gray-700 mt-2">
            <span className="font-semibold font-mono-data">{remaining}</span> left ·{' '}
            <span className="font-semibold font-mono-data">{workingDaysRemaining}</span> working day
            {workingDaysRemaining === 1 ? '' : 's'} ·{' '}
            <span className={`font-semibold ${behind ? 'text-red-700' : 'text-teal-700'}`}>
              ~<span className="font-mono-data">{perDayNeeded}</span>/day
            </span>{' '}
            to finish by {targetDateLabel}
          </p>
        ) : (
          <p className="text-sm text-green-700 font-medium mt-2">
            🎉 All SC parcels are filed.
          </p>
        )}
      </div>
    </div>
  );
}
