import { useMemo, useState } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type SubmittalReview,
  type SubmittalStatusValue,
} from '../lib/sharepoint';
import { formatDateET, formatDateOnly } from '../lib/dates';

// Weekly review statuses. Includes 'Under Review' — a review-only state (not a
// submittal status) for weeks the filing is actively under review.
export const REVIEW_STATUS_OPTIONS: string[] = [
  'Draft', 'Package Mailed (NC)', 'Filed', 'Letter Received - Action Needed',
  'Responded - Awaiting DOR', 'Under Review', 'Approved', 'Denied', 'Withdrawn',
];

// A submittal stops needing weekly reviews once it reaches a closed state.
const CLOSED_STATUSES: SubmittalStatusValue[] = ['Approved', 'Invoiced', 'Paid', 'Denied', 'Withdrawn'];
export const REVIEW_INTERVAL_DAYS = 7;

export function daysSince(iso?: string | null): number | null {
  if (!iso) return null;
  const then = new Date(iso).getTime();
  if (Number.isNaN(then)) return null;
  return Math.floor((Date.now() - then) / 86400000);
}

interface Props {
  submittalId: string;
  currentStatus?: SubmittalStatusValue;
}

export function SubmittalReviewsSection({ submittalId, currentStatus }: Props) {
  const { data, loading, error, refetch } = useSharePointList<SubmittalReview>(
    LIST_NAMES.SubmittalReviews,
    { top: 500 },
  );

  const [status, setStatus] = useState<string>('');
  const [note, setNote] = useState('');
  const [nextAction, setNextAction] = useState('');
  const [eta, setEta] = useState('');
  const [posting, setPosting] = useState(false);
  const [postError, setPostError] = useState<string | null>(null);

  const reviews = useMemo(() => {
    if (!data) return [];
    return data
      .filter((r) => String(r.fields.ReviewSubmittalLookupId ?? '') === String(submittalId))
      .sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime());
  }, [data, submittalId]);

  const closed = currentStatus ? CLOSED_STATUSES.includes(currentStatus) : false;
  const sinceDays = daysSince(reviews[0]?.createdDateTime ?? null);
  const reviewDue = !closed && (sinceDays === null || sinceDays >= REVIEW_INTERVAL_DAYS);
  const effStatus = status || currentStatus || '';

  const handleLog = async () => {
    if (!effStatus) { setPostError('Pick the current status.'); return; }
    setPosting(true);
    setPostError(null);
    try {
      const title = `${effStatus} — ${new Date().toISOString().slice(0, 10)}`.slice(0, 255);
      await createListItem(LIST_NAMES.SubmittalReviews, {
        Title: title,
        ReviewSubmittalLookupId: submittalId,
        ReviewStatus: effStatus,
        ReviewNote: note.trim() || undefined,
        ReviewNextAction: nextAction.trim() || undefined,
        ReviewNextActionETA: eta ? new Date(eta).toISOString() : undefined,
      });
      setNote(''); setNextAction(''); setEta(''); setStatus('');
      await refetch();
    } catch (e) {
      setPostError(e instanceof Error ? e.message : String(e));
    } finally {
      setPosting(false);
    }
  };

  const subtitle = closed
    ? 'Submittal is closed — weekly reviews complete.'
    : sinceDays === null
      ? 'No weekly review logged yet.'
      : `Last reviewed ${sinceDays === 0 ? 'today' : `${sinceDays} day${sinceDays === 1 ? '' : 's'} ago`}.`;

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100 flex items-start justify-between gap-2 flex-wrap">
        <div>
          <h3 className="text-sm font-semibold text-gray-700 flex items-center gap-2">
            Weekly Reviews
            {reviewDue && (
              <span className="inline-block px-2 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800">Review due</span>
            )}
          </h3>
          <p className="text-xs text-gray-500 mt-0.5">{subtitle}</p>
        </div>
      </div>

      {error ? (
        <div className="p-4 text-xs text-red-700">
          Reviews list unavailable — most likely the “Submittal Reviews” list hasn’t been provisioned yet.
          <div className="mt-1 font-mono-data text-red-600">{error.message}</div>
        </div>
      ) : (
        <>
          {/* Log this week's review */}
          {!closed && (
            <div className="p-4 border-b border-gray-100 bg-gray-50/50">
              <div className="flex flex-wrap items-end gap-2">
                <label className="flex flex-col gap-1 text-xs">
                  <span className="font-semibold text-gray-600">Status this week</span>
                  <select
                    value={effStatus}
                    onChange={(e) => setStatus(e.target.value)}
                    disabled={posting}
                    className="border border-gray-300 rounded px-2 py-1 bg-white"
                  >
                    {REVIEW_STATUS_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                  </select>
                </label>
                <label className="flex flex-col gap-1 text-xs">
                  <span className="font-semibold text-gray-600">Next action ETA</span>
                  <input
                    type="date" value={eta} onChange={(e) => setEta(e.target.value)} disabled={posting}
                    className="border border-gray-300 rounded px-2 py-1 font-mono-data"
                  />
                </label>
                <button
                  onClick={handleLog}
                  disabled={posting}
                  className="px-3 py-1.5 rounded-md text-xs font-medium bg-teal-700 hover:bg-teal-900 text-white disabled:opacity-50"
                >
                  {posting ? 'Logging…' : 'Log review'}
                </button>
              </div>
              <textarea
                value={note} onChange={(e) => setNote(e.target.value)} rows={2} disabled={posting}
                placeholder="This week's status / progress note…"
                className="mt-2 w-full border border-gray-300 rounded px-2 py-1 text-sm resize-y disabled:opacity-50"
              />
              <input
                value={nextAction} onChange={(e) => setNextAction(e.target.value)} disabled={posting}
                placeholder="Planned next action…"
                className="mt-2 w-full border border-gray-300 rounded px-2 py-1 text-sm disabled:opacity-50"
              />
              {postError && <p className="text-xs text-red-700 mt-1">{postError}</p>}
            </div>
          )}

          {/* History */}
          {loading ? (
            <div className="p-4 text-xs text-gray-500">Loading reviews…</div>
          ) : reviews.length === 0 ? (
            <div className="p-4 text-xs text-gray-500">No weekly reviews recorded yet.</div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {reviews.map((r) => (
                <li key={r.id} className="px-4 py-3 text-sm">
                  <div className="flex items-center justify-between gap-2 flex-wrap">
                    <span className="inline-block px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-100 text-gray-800">
                      {r.fields.ReviewStatus ?? '—'}
                    </span>
                    <div className="text-xs text-gray-500 font-mono-data text-right">
                      <span>{formatDateET(r.createdDateTime)}</span>
                      <span className="text-gray-400"> · {r.createdBy?.user?.displayName ?? ''}</span>
                    </div>
                  </div>
                  {r.fields.ReviewNote && <p className="mt-1 text-gray-800 whitespace-pre-wrap">{r.fields.ReviewNote}</p>}
                  {(r.fields.ReviewNextAction || r.fields.ReviewNextActionETA) && (
                    <p className="mt-1 text-xs text-gray-600">
                      <span className="text-gray-400">→ </span>
                      {r.fields.ReviewNextAction}
                      {r.fields.ReviewNextActionETA && (
                        <span className="text-gray-400"> · ETA {formatDateOnly(r.fields.ReviewNextActionETA)}</span>
                      )}
                    </p>
                  )}
                </li>
              ))}
            </ul>
          )}
        </>
      )}
    </div>
  );
}
