import { useEffect, useMemo, useState } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Contact,
  type ContactOwnerLink,
  type EmailTemplate,
  type OutstandingItem,
  type Property,
} from '../lib/sharepoint';
import { formatDateOnly, parseDateOnly } from '../lib/dates';
import { sendEmail, applyTemplateVars, type EmailRecipient } from '../lib/email';
import { useSession } from '../lib/session';
import { TEAM_MEMBERS } from '../lib/roleMap';

/**
 * Compose & send an email from inside the app.
 *
 * Flow on Send:
 *   1. Call Graph sendMail as the signed-in user (Mail.Send permission)
 *   2. Create an Owner Communication row with CommType='Email', Outbound,
 *      body in CommNotes, subject as Title
 *   3. Create one CommunicationPropertyLinks row per linked property
 *   4. Create one CommunicationOwnerLinks row per owner derived from the
 *      selected contacts' linked-owner junctions
 *
 * If the email goes through but the auto-log step fails (rare), the user
 * sees a partial-success banner so they know to log it manually.
 */

export interface ComposeEmailModalProps {
  onClose: () => void;
  onSuccess?: () => void;
  /** Pre-select these contacts when opening (used from Contacts row / OwnerDetail). */
  defaultContactIds?: string[];
  /** Pre-select these properties (used from PropertyDetail). */
  defaultPropertyIds?: string[];
  /** Pre-fill subject (used by some entry points). */
  defaultSubject?: string;
  /** Pre-fill body. */
  defaultBody?: string;
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function ComposeEmailModal({
  onClose,
  onSuccess,
  defaultContactIds,
  defaultPropertyIds,
  defaultSubject,
  defaultBody,
}: ComposeEmailModalProps) {
  const { user } = useSession();
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const templates = useSharePointList<EmailTemplate>(LIST_NAMES.EmailTemplates, { top: 500 });
  const contactOwnerLinks = useSharePointList<ContactOwnerLink>(LIST_NAMES.ContactOwnerLinks, { top: 2000 });
  // For {{open_items}} substitution — Outstanding Items assigned to the
  // selected contact(s), optionally filtered to the linked properties.
  const outstandingItems = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });

  const [selectedContactIds, setSelectedContactIds] = useState<Set<string>>(
    new Set(defaultContactIds ?? []),
  );
  const [adHocRecipients, setAdHocRecipients] = useState(''); // comma-separated emails
  const [contactSearch, setContactSearch] = useState('');

  // CC: a Set of team-member emails (quick-pick checkboxes) + a free-text field
  // for any other internal/external addresses. Mirrors the To pattern but
  // doesn't pull from the Contacts list — that's already covered above.
  const [ccTeamEmails, setCcTeamEmails] = useState<Set<string>>(new Set());
  const [ccFreeText, setCcFreeText] = useState('');
  const [showCc, setShowCc] = useState(false);

  const [selectedPropertyIds, setSelectedPropertyIds] = useState<Set<string>>(
    new Set(defaultPropertyIds ?? []),
  );
  const [propertySearch, setPropertySearch] = useState('');

  const [templateId, setTemplateId] = useState('');
  const [subject, setSubject] = useState(defaultSubject ?? '');
  const [body, setBody] = useState(defaultBody ?? '');

  const [sending, setSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [partialError, setPartialError] = useState<string | null>(null);

  // Sort contacts; filter by search
  const sortedContacts = useMemo(
    () => [...(contacts.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    ),
    [contacts.data],
  );
  const filteredContacts = useMemo(() => {
    const q = contactSearch.trim().toLowerCase();
    if (!q) return sortedContacts;
    return sortedContacts.filter((c) => {
      const hay = `${c.fields.Title ?? ''} ${c.fields.ContactEmail ?? ''} ${c.fields.ContactRole ?? ''}`.toLowerCase();
      return hay.includes(q);
    });
  }, [sortedContacts, contactSearch]);

  const sortedProperties = useMemo(
    () => [...(properties.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    ),
    [properties.data],
  );
  const filteredProperties = useMemo(() => {
    const q = propertySearch.trim().toLowerCase();
    if (!q) return sortedProperties;
    return sortedProperties.filter((p) => (p.fields.Title ?? '').toLowerCase().includes(q));
  }, [sortedProperties, propertySearch]);

  const toggleContact = (id: string) =>
    setSelectedContactIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  const toggleProperty = (id: string) =>
    setSelectedPropertyIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });

  // Pick a template — fills subject + body from its values (with vars unresolved
  // so the user can see what'll change at send time).
  const applyTemplate = (id: string) => {
    setTemplateId(id);
    if (!id) return;
    const tmpl = templates.data?.find((t) => String(t.id) === id);
    if (!tmpl) return;
    setSubject(tmpl.fields.TemplateSubject ?? '');
    setBody(tmpl.fields.TemplateBody ?? '');
  };

  // Build template context for variable substitution at send time.
  // When multiple contacts/properties are selected, {{contact}} = first,
  // {{properties}} = joined list. {{open_items}} pulls Outstanding Items
  // assigned to the selected contacts (matched by name OR email,
  // case-insensitive), filtered to the linked properties when any are
  // chosen.
  const templateContext = useMemo(() => {
    const selectedContacts = (contacts.data ?? []).filter((c) =>
      selectedContactIds.has(String(c.id)),
    );
    const selectedProps = (properties.data ?? []).filter((p) =>
      selectedPropertyIds.has(String(p.id)),
    );
    const primaryContact = selectedContacts[0];
    const primaryProp = selectedProps[0];

    // Build the set of AssignedTo match keys for the selected contacts
    const assigneeKeys = new Set<string>();
    for (const c of selectedContacts) {
      const name = (c.fields.Title ?? '').trim().toLowerCase();
      const email = (c.fields.ContactEmail ?? '').trim().toLowerCase();
      if (name) assigneeKeys.add(name);
      if (email) assigneeKeys.add(email);
    }

    // Only consider properties the email is tagged to; if none selected,
    // include items across every property the recipients have items on.
    const propertyFilter = selectedPropertyIds.size > 0 ? selectedPropertyIds : null;

    const isClosed = (s: string | undefined) =>
      s === 'Done' || s === 'Received' || s === 'Not Applicable';

    const propsById = new Map<string, Property>();
    (properties.data ?? []).forEach((p) => propsById.set(String(p.id), p));

    const matchingItems = assigneeKeys.size === 0
      ? []
      : (outstandingItems.data ?? []).filter((item) => {
          if (isClosed(item.fields.ItemStatus)) return false;
          const a = (item.fields.AssignedTo ?? '').trim().toLowerCase();
          if (!a || !assigneeKeys.has(a)) return false;
          if (propertyFilter) {
            const pid = String(item.fields.PropertyLookupId ?? '');
            if (!propertyFilter.has(pid)) return false;
          }
          return true;
        }).sort((a, b) => {
          // Oldest due first; no-date items go to the bottom
          const ad = a.fields.DueDate ? parseDateOnly(a.fields.DueDate)?.getTime() ?? Number.MAX_SAFE_INTEGER : Number.MAX_SAFE_INTEGER;
          const bd = b.fields.DueDate ? parseDateOnly(b.fields.DueDate)?.getTime() ?? Number.MAX_SAFE_INTEGER : Number.MAX_SAFE_INTEGER;
          return ad - bd;
        });

    const openItemsList = matchingItems.length === 0
      ? '(no pending items for this recipient)'
      : matchingItems
          .map((item) => {
            const prop = item.fields.PropertyLookupId
              ? propsById.get(String(item.fields.PropertyLookupId))?.fields.Title
              : null;
            const due = item.fields.DueDate ? ` (due ${formatDateOnly(item.fields.DueDate)})` : '';
            const propPart = prop ? ` — ${prop}` : '';
            return `  • ${item.fields.Title ?? 'Untitled'}${propPart}${due}`;
          })
          .join('\n');

    return {
      contactName: primaryContact?.fields.Title,
      contactEmail: primaryContact?.fields.ContactEmail,
      propertyName: primaryProp?.fields.Title,
      propertiesList: selectedProps.map((p) => p.fields.Title).filter(Boolean).join(', ') || undefined,
      userName: user?.name,
      userEmail: user?.email,
      openItemsList,
    };
  }, [
    contacts.data,
    properties.data,
    outstandingItems.data,
    selectedContactIds,
    selectedPropertyIds,
    user?.name,
    user?.email,
  ]);

  const resolvedSubject = useMemo(() => applyTemplateVars(subject, templateContext), [subject, templateContext]);
  const resolvedBody = useMemo(() => applyTemplateVars(body, templateContext), [body, templateContext]);

  // Build recipient list: linked contacts + free-text ad hoc addresses
  const recipients = useMemo<EmailRecipient[]>(() => {
    const list: EmailRecipient[] = [];
    const seen = new Set<string>();
    for (const c of contacts.data ?? []) {
      if (!selectedContactIds.has(String(c.id))) continue;
      const addr = (c.fields.ContactEmail ?? '').trim();
      if (!addr || seen.has(addr.toLowerCase())) continue;
      seen.add(addr.toLowerCase());
      list.push({ address: addr, name: c.fields.Title });
    }
    for (const raw of adHocRecipients.split(',')) {
      const addr = raw.trim();
      if (!addr || seen.has(addr.toLowerCase())) continue;
      seen.add(addr.toLowerCase());
      list.push({ address: addr });
    }
    return list;
  }, [contacts.data, selectedContactIds, adHocRecipients]);

  const recipientErrors = useMemo(() => {
    const out: string[] = [];
    for (const r of recipients) {
      if (!EMAIL_RE.test(r.address)) out.push(r.address);
    }
    return out;
  }, [recipients]);

  // CC recipients = checked team members + free-text additions
  const ccRecipients = useMemo<EmailRecipient[]>(() => {
    const list: EmailRecipient[] = [];
    const seen = new Set<string>();
    // De-dupe against the To list so the same person isn't both To and CC
    for (const r of recipients) seen.add(r.address.toLowerCase());
    for (const m of TEAM_MEMBERS) {
      if (!ccTeamEmails.has(m.email)) continue;
      if (seen.has(m.email.toLowerCase())) continue;
      seen.add(m.email.toLowerCase());
      list.push({ address: m.email, name: m.name });
    }
    for (const raw of ccFreeText.split(',')) {
      const addr = raw.trim();
      if (!addr || seen.has(addr.toLowerCase())) continue;
      seen.add(addr.toLowerCase());
      list.push({ address: addr });
    }
    return list;
  }, [ccTeamEmails, ccFreeText, recipients]);

  const ccErrors = useMemo(
    () => ccRecipients.filter((r) => !EMAIL_RE.test(r.address)).map((r) => r.address),
    [ccRecipients],
  );

  const toggleCcTeam = (email: string) =>
    setCcTeamEmails((prev) => {
      const next = new Set(prev);
      if (next.has(email)) next.delete(email); else next.add(email);
      return next;
    });

  // Owners linked to the selected contacts — used for auto-log junction rows
  const linkedOwnerIds = useMemo(() => {
    const ids = new Set<string>();
    for (const c of contacts.data ?? []) {
      if (!selectedContactIds.has(String(c.id))) continue;
      // Legacy single-owner field
      if (c.fields.ContactOwnerLookupId) ids.add(String(c.fields.ContactOwnerLookupId));
    }
    for (const link of contactOwnerLinks.data ?? []) {
      const cid = String(link.fields.ContactLookupId ?? '');
      if (!cid || !selectedContactIds.has(cid)) continue;
      if (link.fields.OwnerLookupId) ids.add(String(link.fields.OwnerLookupId));
    }
    return ids;
  }, [contacts.data, contactOwnerLinks.data, selectedContactIds]);

  // Auto-prefill the first selected property when arriving from a property page
  useEffect(() => {
    if (defaultPropertyIds && defaultPropertyIds.length > 0) {
      setSelectedPropertyIds(new Set(defaultPropertyIds));
    }
  }, [defaultPropertyIds]);

  const handleSend = async () => {
    setError(null);
    setPartialError(null);

    if (recipients.length === 0) {
      setError('Add at least one recipient.');
      return;
    }
    if (recipientErrors.length > 0) {
      setError(`Invalid recipient email${recipientErrors.length === 1 ? '' : 's'}: ${recipientErrors.join(', ')}`);
      return;
    }
    if (ccErrors.length > 0) {
      setError(`Invalid CC email${ccErrors.length === 1 ? '' : 's'}: ${ccErrors.join(', ')}`);
      return;
    }
    if (!resolvedSubject.trim()) {
      setError('Subject is required.');
      return;
    }

    setSending(true);
    try {
      // 1. Send via Graph
      await sendEmail({
        to: recipients,
        cc: ccRecipients.length > 0 ? ccRecipients : undefined,
        subject: resolvedSubject,
        bodyText: resolvedBody,
      });

      // 2. Auto-log as an Owner Communication
      try {
        const toLine = recipients.map((r) => r.name ? `${r.name} <${r.address}>` : r.address).join(', ');
        const ccLine = ccRecipients.length > 0
          ? ' · CC: ' + ccRecipients.map((r) => r.name ? `${r.name} <${r.address}>` : r.address).join(', ')
          : '';
        const recipientLine = `To: ${toLine}${ccLine}`;
        const propIds = Array.from(selectedPropertyIds);
        const ownerIds = Array.from(linkedOwnerIds);
        const commPayload: Record<string, unknown> = {
          Title: resolvedSubject,
          CommType: 'Email',
          CommDirection: 'Outbound',
          CommDate: new Date().toISOString(),
          CommStatus: 'Closed',
          CommParticipants: recipientLine,
          CommNotes: resolvedBody,
        };
        if (propIds[0]) commPayload.CommPropertyLookupId = propIds[0];
        if (ownerIds[0]) commPayload.CommOwnerLookupId = ownerIds[0];
        const comm = await createListItem<{ id: string }>(LIST_NAMES.Communications, commPayload);

        // Junction rows for properties + owners
        for (const pid of propIds) {
          try {
            await createListItem(LIST_NAMES.CommunicationPropertyLinks, {
              Title: `Comm ${comm.id} ↔ Property ${pid}`,
              CommLookupId: Number(comm.id),
              PropertyLookupId: Number(pid),
            });
          } catch {
            /* silent — primary single field still captures one */
          }
        }
        for (const oid of ownerIds) {
          try {
            await createListItem(LIST_NAMES.CommunicationOwnerLinks, {
              Title: `Comm ${comm.id} ↔ Owner ${oid}`,
              CommLookupId: Number(comm.id),
              OwnerLookupId: Number(oid),
            });
          } catch {
            /* silent */
          }
        }
      } catch (logErr) {
        setPartialError(
          'Email sent, but auto-logging to Owner Communications failed: ' +
          (logErr instanceof Error ? logErr.message : String(logErr)) +
          '. Log it manually if you want a record.',
        );
        setTimeout(() => onSuccess?.(), 1500);
        return;
      }

      onSuccess?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSending(false);
    }
  };

  const inputClass =
    'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500';

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-3xl w-full p-5 my-8">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Compose Email</h3>
        <p className="text-sm text-gray-600 mb-4">
          Sends as <strong>{user?.email ?? 'you'}</strong> via your Outlook account.
          A copy lands in your Sent folder, and we'll log it to Owner Communications
          tagged to the selected contacts + properties.
        </p>

        <div className="space-y-3">
          <Field label={`Recipients (${recipients.length})`} required>
            <input
              type="text"
              value={contactSearch}
              onChange={(e) => setContactSearch(e.target.value)}
              placeholder="Search contacts by name, email, or role…"
              disabled={sending}
              className={inputClass + ' mb-1'}
            />
            <div className="border border-gray-300 rounded max-h-32 overflow-y-auto bg-white">
              {filteredContacts.length === 0 ? (
                <div className="px-3 py-2 text-[11px] text-gray-500 italic">
                  {contacts.loading ? 'Loading contacts…' : 'No contacts match your search.'}
                </div>
              ) : (
                filteredContacts.map((c) => (
                  <label key={c.id} className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs">
                    <input
                      type="checkbox"
                      checked={selectedContactIds.has(String(c.id))}
                      onChange={() => toggleContact(String(c.id))}
                      disabled={sending}
                    />
                    <span className="flex-1 truncate">{c.fields.Title}</span>
                    {c.fields.ContactEmail && (
                      <span className="text-[10px] text-gray-500 font-mono-data flex-shrink-0">
                        {c.fields.ContactEmail}
                      </span>
                    )}
                    {c.fields.ContactRole && (
                      <span className="text-[10px] text-gray-500 flex-shrink-0">· {c.fields.ContactRole}</span>
                    )}
                  </label>
                ))
              )}
            </div>
            <input
              type="text"
              value={adHocRecipients}
              onChange={(e) => setAdHocRecipients(e.target.value)}
              placeholder="Or add emails manually, comma-separated (one-off addresses)"
              disabled={sending}
              className={inputClass + ' mt-1'}
            />
            {recipientErrors.length > 0 && (
              <p className="text-[11px] text-error mt-1">
                Invalid: {recipientErrors.join(', ')}
              </p>
            )}
          </Field>

          {/* CC — team-member quick-pick + free-text fallback */}
          {!showCc && ccRecipients.length === 0 ? (
            <button
              type="button"
              onClick={() => setShowCc(true)}
              className="text-xs text-teal-700 hover:text-teal-900 font-medium underline self-start"
            >
              + Add CC
            </button>
          ) : (
            <Field label={`CC (${ccRecipients.length})`}>
              <div className="border border-gray-300 rounded bg-white p-2 space-y-1">
                <div className="text-[10px] font-semibold text-gray-500 uppercase tracking-wider">Team members</div>
                <div className="flex flex-wrap gap-2">
                  {TEAM_MEMBERS.map((m) => (
                    <label key={m.email} className="inline-flex items-center gap-1.5 text-xs cursor-pointer">
                      <input
                        type="checkbox"
                        checked={ccTeamEmails.has(m.email)}
                        onChange={() => toggleCcTeam(m.email)}
                        disabled={sending}
                      />
                      <span className="text-gray-800">{m.name}</span>
                      <span className="text-[10px] text-gray-500 font-mono-data">{m.email}</span>
                    </label>
                  ))}
                </div>
              </div>
              <input
                type="text"
                value={ccFreeText}
                onChange={(e) => setCcFreeText(e.target.value)}
                placeholder="Other CC addresses, comma-separated"
                disabled={sending}
                className={inputClass + ' mt-1'}
              />
              {ccErrors.length > 0 && (
                <p className="text-[11px] text-error mt-1">
                  Invalid CC: {ccErrors.join(', ')}
                </p>
              )}
              <p className="text-[11px] text-gray-500 mt-1">
                CC'd folks are also captured in the auto-logged Owner Communication participants line.
              </p>
            </Field>
          )}

          <Field label={`Linked Properties (${selectedPropertyIds.size})`}>
            <input
              type="text"
              value={propertySearch}
              onChange={(e) => setPropertySearch(e.target.value)}
              placeholder="Search properties to link…"
              disabled={sending}
              className={inputClass + ' mb-1'}
            />
            <div className="border border-gray-300 rounded max-h-32 overflow-y-auto bg-white">
              {filteredProperties.length === 0 ? (
                <div className="px-3 py-2 text-[11px] text-gray-500 italic">No properties match.</div>
              ) : (
                filteredProperties.map((p) => (
                  <label key={p.id} className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs">
                    <input
                      type="checkbox"
                      checked={selectedPropertyIds.has(String(p.id))}
                      onChange={() => toggleProperty(String(p.id))}
                      disabled={sending}
                    />
                    <span className="flex-1 truncate">{p.fields.Title}</span>
                    {p.fields.cahpState && (
                      <span className="text-[10px] text-gray-500 flex-shrink-0">{p.fields.cahpState}</span>
                    )}
                  </label>
                ))
              )}
            </div>
            <p className="text-[11px] text-gray-500 mt-1">
              The auto-logged communication is tagged to every linked property + every Owner entity the selected contacts represent.
            </p>
          </Field>

          {templates.data && templates.data.length > 0 && (
            <Field label="Template (optional)">
              <select
                value={templateId}
                onChange={(e) => applyTemplate(e.target.value)}
                disabled={sending}
                className={inputClass + ' bg-white'}
              >
                <option value="">— start from blank —</option>
                {[...templates.data]
                  .sort((a, b) => (a.fields.TemplateSortOrder ?? 0) - (b.fields.TemplateSortOrder ?? 0))
                  .map((t) => (
                    <option key={t.id} value={t.id}>{t.fields.Title}</option>
                  ))}
              </select>
              <p className="text-[11px] text-gray-500 mt-0.5">
                Variables like <code className="bg-gray-100 px-1 rounded">{'{{contact}}'}</code> get substituted at send time using the first selected recipient + property.
              </p>
            </Field>
          )}

          <Field label="Subject" required>
            <input
              type="text"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              disabled={sending}
              placeholder="What's this email about?"
              className={inputClass}
            />
            {subject !== resolvedSubject && (
              <p className="text-[11px] text-gray-500 mt-1">
                Preview: <span className="font-mono-data">{resolvedSubject}</span>
              </p>
            )}
          </Field>

          <Field label="Body">
            <textarea
              value={body}
              onChange={(e) => setBody(e.target.value)}
              disabled={sending}
              rows={10}
              placeholder="Your message…"
              className={inputClass + ' font-mono text-xs resize-y'}
            />
            {body !== resolvedBody && (
              <details className="mt-1 text-[11px] text-gray-500">
                <summary className="cursor-pointer hover:text-gray-700">Preview with variables resolved</summary>
                <pre className="mt-1 p-2 bg-gray-50 border border-gray-200 rounded whitespace-pre-wrap font-mono text-[11px]">{resolvedBody}</pre>
              </details>
            )}
          </Field>
        </div>

        {error && (
          <div className="mt-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">{error}</div>
        )}
        {partialError && (
          <div className="mt-3 bg-amber-50 border border-amber-200 rounded-md p-2 text-xs text-amber-900">{partialError}</div>
        )}

        <div className="flex items-center justify-end gap-2 mt-5">
          <button
            onClick={onClose}
            disabled={sending}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleSend}
            disabled={sending || recipients.length === 0 || !resolvedSubject.trim() || recipientErrors.length > 0 || ccErrors.length > 0}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:cursor-not-allowed"
          >
            {sending && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {sending ? 'Sending…' : `Send to ${recipients.length}`}
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, children }: { label: string; required?: boolean; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
        {label}{required && <span className="text-error ml-0.5">*</span>}
      </label>
      {children}
    </div>
  );
}
