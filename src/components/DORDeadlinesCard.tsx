import { useMemo } from 'react';
import { Link } from 'react-router-dom';
import { Icon } from './ui/Icon';
import { EASTERN_TZ } from '../lib/dates';
import { selectDORDeadlines, type DORDeadlineKind } from '../lib/dor-clock';
import type { Submittal, Property } from '../lib/sharepoint';

/**
 * "DOR Deadlines" card — surfaces every submittal with a live DOR clock running:
 *   - Awaiting DOR response (their court, ~12 weeks from filing/our response)
 *   - RFI response owed to DOR (our court, 30 days from receipt)
 * Overdue items bubble to the top and are flagged red. Renders nothing when
 * there are no live clocks.
 */

const KIND_META: Record<DORDeadlineKind, { label: string; verb: string; badge: string }> = {
  'awaiting-dor': {
    label: 'Awaiting DOR',
    verb: 'DOR response expected',
    badge: 'bg-purple-100 text-purple-800',
  },
  'rfi-response': {
    label: 'RFI — our response',
    verb: 'Our response due',
    badge: 'bg-amber-100 text-amber-800',
  },
};

export function DORDeadlinesCard({
  submittals,
  propertiesById,
  headerLink,
  limit = 6,
}: {
  submittals: Submittal[] | undefined | null;
  propertiesById: Map<string, Property>;
  /** Optional "View all →" target (e.g. the Submittals page from My Day). */
  headerLink?: string;
  limit?: number;
}) {
  const entries = useMemo(() => selectDORDeadlines(submittals), [submittals]);

  if (entries.length === 0) return null;

  const overdueCount = entries.filter((e) => e.isOverdue).length;
  const shown = entries.slice(0, limit);
  const hidden = entries.length - shown.length;

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6">
      <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
        <div>
          <h2 className="text-base font-semibold text-teal-700 flex items-center gap-2">
            <Icon name="calendar" size={16} />
            DOR Deadlines
            {overdueCount > 0 && (
              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-red-100 text-error">
                {overdueCount} overdue
              </span>
            )}
          </h2>
          <p className="text-xs text-gray-500 mt-0.5">
            {entries.length} live clock{entries.length === 1 ? '' : 's'} · RFI replies due in 30 days · DOR responses expected in ~12 weeks
          </p>
        </div>
        {headerLink && (
          <Link to={headerLink} className="text-xs text-teal-700 hover:text-teal-900 font-medium">
            View all →
          </Link>
        )}
      </div>
      <div className="divide-y divide-gray-100">
        {shown.map((e) => {
          const s = e.submittal;
          const meta = KIND_META[e.kind];
          const property = s.fields.PropertyLookupId
            ? propertiesById.get(String(s.fields.PropertyLookupId))
            : null;
          return (
            <Link
              key={s.id}
              to={`/submittals/${s.id}`}
              className="px-5 py-3 flex items-center gap-3 hover:bg-gray-50 transition-colors"
            >
              <div className="flex-1 min-w-0">
                <div className="text-sm font-medium text-gray-900 truncate">
                  {e.isOverdue && <span className="text-error mr-1">⚠</span>}
                  {property?.fields.Title ?? s.fields.Title ?? '(unlinked submittal)'}
                </div>
                <div className="text-xs text-gray-500 mt-0.5 truncate flex items-center gap-1.5">
                  <span className={`px-1.5 py-0.5 rounded text-[10px] font-semibold ${meta.badge}`}>{meta.label}</span>
                  <span className="truncate">{s.fields.Title}</span>
                  {s.fields.cahpTaxYear && (
                    <span className="font-mono-data text-teal-700">TY{s.fields.cahpTaxYear}</span>
                  )}
                </div>
              </div>
              <div className="text-right flex-shrink-0 w-24">
                {e.due ? (
                  <>
                    <div className={`text-xs font-mono-data ${e.isOverdue ? 'text-error font-bold' : 'text-gray-700'}`}>
                      {e.due.toLocaleDateString('en-US', { timeZone: EASTERN_TZ, month: 'short', day: 'numeric', year: 'numeric' })}
                    </div>
                    <div className="text-[10px] text-gray-500">
                      {meta.verb} ·{' '}
                      {e.isOverdue
                        ? `${Math.abs(e.daysOut!)}d overdue`
                        : e.daysOut === 0
                          ? 'today'
                          : `in ${e.daysOut}d`}
                    </div>
                  </>
                ) : (
                  <span className="text-[10px] text-gray-400">no date set</span>
                )}
              </div>
            </Link>
          );
        })}
      </div>
      {hidden > 0 && (
        <div className="px-5 py-2 text-[11px] text-gray-400 italic border-t border-gray-100">
          …and {hidden} more
        </div>
      )}
    </div>
  );
}
