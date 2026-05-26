import { useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type OwnerCommunication,
  type CommunicationPropertyLink,
  type CommunicationOwnerLink,
} from '../lib/sharepoint';
import { formatDateOnly } from '../lib/dates';
import { Icon } from './ui/Icon';

/**
 * Displays the Owner Communications list filtered to a given property or
 * owner — so emails sent via the Compose modal (which auto-log to that list
 * along with junction rows) show up on the entity's detail page for quick
 * reference.
 *
 * Matches via:
 *   - Direct field (CommPropertyLookupId / CommOwnerLookupId) — the "primary"
 *     legacy linkage, populated on every comm.
 *   - Junction list (CommunicationPropertyLinks / CommunicationOwnerLinks) —
 *     so multi-property emails surface on every linked property.
 */
export function LoggedCommunicationsSection({
  propertyId,
  ownerId,
  title = 'Communications Log',
  subtitle,
}: {
  propertyId?: string;
  ownerId?: string;
  title?: string;
  subtitle?: string;
}) {
  const navigate = useNavigate();
  const comms = useSharePointList<OwnerCommunication>(LIST_NAMES.Communications, { top: 500 });
  const propLinks = useSharePointList<CommunicationPropertyLink>(LIST_NAMES.CommunicationPropertyLinks, { top: 2000 });
  const ownerLinks = useSharePointList<CommunicationOwnerLink>(LIST_NAMES.CommunicationOwnerLinks, { top: 2000 });

  const matched = useMemo(() => {
    if (!comms.data) return [];
    const linkedIds = new Set<string>();
    if (propertyId) {
      for (const link of propLinks.data ?? []) {
        if (String(link.fields.PropertyLookupId ?? '') !== String(propertyId)) continue;
        if (link.fields.CommLookupId) linkedIds.add(String(link.fields.CommLookupId));
      }
    }
    if (ownerId) {
      for (const link of ownerLinks.data ?? []) {
        if (String(link.fields.OwnerLookupId ?? '') !== String(ownerId)) continue;
        if (link.fields.CommLookupId) linkedIds.add(String(link.fields.CommLookupId));
      }
    }
    return comms.data
      .filter((c) => {
        if (propertyId && String(c.fields.CommPropertyLookupId ?? '') === String(propertyId)) return true;
        if (ownerId && String(c.fields.CommOwnerLookupId ?? '') === String(ownerId)) return true;
        return linkedIds.has(String(c.id));
      })
      .sort((a, b) => {
        const ad = a.fields.CommDate ? new Date(a.fields.CommDate).getTime() : 0;
        const bd = b.fields.CommDate ? new Date(b.fields.CommDate).getTime() : 0;
        return bd - ad;
      });
  }, [comms.data, propLinks.data, ownerLinks.data, propertyId, ownerId]);

  const loading = comms.loading || propLinks.loading || ownerLinks.loading;

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100">
        <h3 className="text-sm font-semibold text-gray-700">{title}</h3>
        {subtitle && <p className="text-xs text-gray-500 mt-0.5">{subtitle}</p>}
        {!subtitle && (
          <p className="text-xs text-gray-500 mt-0.5">
            Emails and DOR notes auto-logged when sent via the app, plus any items you added manually.
            {matched.length > 0 && <> {matched.length} on file.</>}
          </p>
        )}
      </div>
      {loading ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500">
          <div className="inline-flex items-center gap-2">
            <div className="w-3 h-3 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
            Loading communications…
          </div>
        </div>
      ) : matched.length === 0 ? (
        <div className="px-4 py-6 text-center text-sm text-gray-500 italic">
          No communications logged yet. When you send an email from the app it lands here automatically.
        </div>
      ) : (
        <ul className="divide-y divide-gray-100">
          {matched.map((c) => {
            const f = c.fields;
            const inbound = f.CommDirection === 'Inbound';
            return (
              <li key={c.id}>
                <button
                  type="button"
                  onClick={() => navigate(`/communications/${c.id}`)}
                  className="w-full text-left px-4 py-2.5 hover:bg-gray-50 flex items-start gap-3"
                >
                  <Icon
                    name={inbound ? 'inbox' : 'mail'}
                    size={14}
                    className={`mt-0.5 flex-shrink-0 ${inbound ? 'text-blue-600' : 'text-teal-700'}`}
                  />
                  <div className="flex-1 min-w-0">
                    <div className="flex items-baseline justify-between gap-2">
                      <span className="text-sm font-medium text-gray-900 truncate">{f.Title || '(no subject)'}</span>
                      <span className="text-[11px] text-gray-500 font-mono-data flex-shrink-0">
                        {f.CommDate ? formatDateOnly(f.CommDate) : '—'}
                      </span>
                    </div>
                    <div className="text-[11px] text-gray-500 mt-0.5 truncate">
                      {f.CommType ?? 'Email'} · {f.CommDirection ?? 'Outbound'}
                      {f.CommParticipants && <> · {f.CommParticipants}</>}
                    </div>
                  </div>
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}
