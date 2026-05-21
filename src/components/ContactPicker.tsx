import { useState, useMemo, useRef, useEffect } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Contact,
  type ContactFields,
  type ContactRole,
  type ContactOwnerLink,
  type Owner,
} from '../lib/sharepoint';
import { Icon } from './ui/Icon';

/**
 * ContactPicker — autocomplete dropdown for selecting an existing Contact, with
 * an inline "Add new" option that opens the NewContactModal and selects the
 * newly created contact on success.
 *
 * Stores the contact's listItem ID via `onChange(contactId | undefined)`.
 *
 * Used by Property Overview to pick the property's owner-side point of contact.
 */
export interface ContactPickerProps {
  /** Currently selected contact ID (or undefined for none). */
  value: string | undefined;
  /** Called when the user picks a different contact (or clears it). */
  onChange: (contactId: string | undefined) => void;
  /** Disable interaction. */
  disabled?: boolean;
  /** Optional: pre-fill this Owner lookup when adding a new contact. */
  defaultOwnerId?: string;
  /** Hide the create-new affordance (read-only-ish picker). */
  hideCreate?: boolean;
  className?: string;
}

const CONTACT_ROLES: ContactRole[] = [
  'Property Owner',
  'Sponsor',
  'Attorney',
  'Accountant',
  'Property Manager',
  'Vendor',
  'Lender',
  'Other',
];

export function ContactPicker({
  value,
  onChange,
  disabled,
  defaultOwnerId,
  hideCreate,
  className = '',
}: ContactPickerProps) {
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownerLinks = useSharePointList<ContactOwnerLink>(LIST_NAMES.ContactOwnerLinks, { top: 2000 });
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  const [creating, setCreating] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Close the dropdown when clicking outside
  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  const ownersById = useMemo(() => {
    const m = new Map<string, Owner>();
    (owners.data ?? []).forEach((o) => m.set(String(o.id), o));
    return m;
  }, [owners.data]);

  // contactId → ordered list of owner names linked to that contact (junction).
  // Includes the legacy single ContactOwnerLookupId too for backward compat.
  const ownerNamesByContact = useMemo(() => {
    const map = new Map<string, string[]>();
    (ownerLinks.data ?? []).forEach((l) => {
      const cId = l.fields.ContactLookupId ? String(l.fields.ContactLookupId) : '';
      const oId = l.fields.OwnerLookupId ? String(l.fields.OwnerLookupId) : '';
      if (!cId || !oId) return;
      const name = ownersById.get(oId)?.fields.Title;
      if (!name) return;
      if (!map.has(cId)) map.set(cId, []);
      const list = map.get(cId)!;
      if (!list.includes(name)) list.push(name);
    });
    // Layer in legacy single link as a fallback
    (contacts.data ?? []).forEach((c) => {
      const legacyId = c.fields.ContactOwnerLookupId
        ? String(c.fields.ContactOwnerLookupId)
        : '';
      if (!legacyId) return;
      const name = ownersById.get(legacyId)?.fields.Title;
      if (!name) return;
      const cId = String(c.id);
      if (!map.has(cId)) map.set(cId, []);
      const list = map.get(cId)!;
      if (!list.includes(name)) list.push(name);
    });
    return map;
  }, [ownerLinks.data, ownersById, contacts.data]);

  const selected = useMemo(() => {
    if (!value) return undefined;
    return (contacts.data ?? []).find((c) => String(c.id) === String(value));
  }, [contacts.data, value]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    const all = [...(contacts.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    );
    if (!q) return all;
    return all.filter((c) => {
      const haystack = `${c.fields.Title ?? ''} ${c.fields.ContactEmail ?? ''} ${c.fields.ContactRole ?? ''}`.toLowerCase();
      return haystack.includes(q);
    });
  }, [contacts.data, query]);

  const handleSelect = (contactId: string) => {
    onChange(contactId);
    setQuery('');
    setOpen(false);
  };

  const handleClear = () => {
    onChange(undefined);
    setQuery('');
  };

  return (
    <div ref={containerRef} className={`relative ${className}`}>
      {selected ? (
        <div className="flex items-center justify-between gap-2 px-2 py-1.5 border border-gray-300 rounded bg-white">
          <div className="flex-1 min-w-0">
            <div className="text-sm font-medium text-gray-900 truncate">{selected.fields.Title}</div>
            <div className="text-[11px] text-gray-500 truncate">
              {selected.fields.ContactEmail || '(no email)'}
              {selected.fields.ContactRole && <span className="ml-1">· {selected.fields.ContactRole}</span>}
              {(ownerNamesByContact.get(String(selected.id)) ?? []).length > 0 && (
                <span className="ml-1">
                  · {(ownerNamesByContact.get(String(selected.id)) ?? []).join(', ')}
                </span>
              )}
            </div>
          </div>
          {!disabled && (
            <button
              type="button"
              onClick={handleClear}
              className="text-[11px] text-gray-500 hover:text-red-700 px-1.5 py-0.5 rounded hover:bg-red-50 flex-shrink-0"
              title="Clear contact"
            >
              Clear
            </button>
          )}
        </div>
      ) : (
        <input
          type="text"
          value={query}
          onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          disabled={disabled}
          placeholder="Search contacts by name or email…"
          className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 disabled:bg-gray-50"
        />
      )}

      {open && !selected && (
        <div className="absolute z-30 mt-1 left-0 right-0 bg-white border border-gray-200 rounded-md shadow-lg max-h-60 overflow-y-auto">
          {contacts.loading ? (
            <div className="px-3 py-3 text-xs text-gray-500">Loading contacts…</div>
          ) : filtered.length === 0 ? (
            <div className="px-3 py-3 text-xs text-gray-500">
              {query ? `No contacts match "${query}".` : 'No contacts in the system yet.'}
            </div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {filtered.slice(0, 20).map((c) => {
                const ownerNames = ownerNamesByContact.get(String(c.id)) ?? [];
                return (
                  <li key={c.id}>
                    <button
                      type="button"
                      onClick={() => handleSelect(String(c.id))}
                      className="w-full text-left px-3 py-2 hover:bg-teal-50"
                    >
                      <div className="text-sm font-medium text-gray-900">{c.fields.Title}</div>
                      <div className="text-[11px] text-gray-500">
                        {c.fields.ContactEmail || '(no email)'}
                        {c.fields.ContactRole && <span className="ml-1">· {c.fields.ContactRole}</span>}
                        {ownerNames.length > 0 && (
                          <span className="ml-1">· {ownerNames.join(', ')}</span>
                        )}
                      </div>
                    </button>
                  </li>
                );
              })}
            </ul>
          )}
          {!hideCreate && (
            <button
              type="button"
              onClick={() => { setOpen(false); setCreating(true); }}
              className="w-full text-left px-3 py-2 border-t border-gray-200 bg-gray-50 hover:bg-teal-50 text-xs text-teal-700 font-medium inline-flex items-center gap-1.5"
            >
              <Icon name="plus" size={11} />
              Add new contact{query.trim() && <span className="text-gray-500"> · prefill name "{query.trim()}"</span>}
            </button>
          )}
        </div>
      )}

      {creating && (
        <NewContactModal
          defaultName={query.trim() || ''}
          defaultOwnerId={defaultOwnerId}
          onClose={() => setCreating(false)}
          onSaved={(newId) => {
            contacts.refetch?.();
            setCreating(false);
            onChange(newId);
            setQuery('');
          }}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// New Contact modal — also exported so other places can open it directly
// ---------------------------------------------------------------------------

export interface NewContactModalProps {
  defaultName?: string;
  defaultOwnerId?: string;
  onClose: () => void;
  onSaved: (newContactId: string) => void;
}

export function NewContactModal({
  defaultName,
  defaultOwnerId,
  onClose,
  onSaved,
}: NewContactModalProps) {
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const [title, setTitle] = useState(defaultName ?? '');
  const [email, setEmail] = useState('');
  const [phone, setPhone] = useState('');
  const [role, setRole] = useState<ContactRole | ''>('');
  // Multi-owner linkage — one contact can represent many Owner entities
  const [selectedOwnerIds, setSelectedOwnerIds] = useState<Set<string>>(
    new Set(defaultOwnerId ? [defaultOwnerId] : []),
  );
  const [ownerSearch, setOwnerSearch] = useState('');
  const [notes, setNotes] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const sortedOwners = useMemo(
    () => [...(owners.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    ),
    [owners.data],
  );

  const filteredOwners = useMemo(() => {
    const q = ownerSearch.trim().toLowerCase();
    if (!q) return sortedOwners;
    return sortedOwners.filter((o) => (o.fields.Title ?? '').toLowerCase().includes(q));
  }, [sortedOwners, ownerSearch]);

  const toggleOwner = (ownerId: string) => {
    setSelectedOwnerIds((prev) => {
      const next = new Set(prev);
      if (next.has(ownerId)) next.delete(ownerId);
      else next.add(ownerId);
      return next;
    });
  };

  const handleSave = async () => {
    if (!title.trim()) {
      setError('Contact name is required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      // Save the primary single-owner field as the FIRST selected owner so the
      // legacy ContactOwner column on the SharePoint list keeps a meaningful
      // value (default views, exports). Junction is the source of truth.
      const primaryOwnerId = selectedOwnerIds.size > 0
        ? Array.from(selectedOwnerIds)[0]
        : undefined;

      const payload: ContactFields & { Title: string } = {
        Title: title.trim(),
        ContactEmail: email.trim() || undefined,
        ContactPhone: phone.trim() || undefined,
        ContactRole: role || undefined,
        ContactOwnerLookupId: primaryOwnerId,
        ContactNotes: notes.trim() || undefined,
      };
      const created = await createListItem<Contact>(LIST_NAMES.Contacts, payload as unknown as Record<string, unknown>);
      const newContactId = String(created.id);

      // Create one junction row per selected owner
      for (const ownerId of selectedOwnerIds) {
        await createListItem(LIST_NAMES.ContactOwnerLinks, {
          Title: `Contact ${newContactId} ↔ Owner ${ownerId}`,
          ContactLookupId: Number(newContactId),
          OwnerLookupId: Number(ownerId),
        });
      }

      onSaved(newContactId);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !saving) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-md w-full max-h-[90vh] overflow-y-auto">
        <div className="px-5 py-3 border-b border-gray-200">
          <h2 className="text-base font-bold text-teal-700">Add Contact</h2>
          <p className="text-xs text-gray-500 mt-0.5">
            People we communicate with about properties — assignees, owners, attorneys, vendors.
          </p>
        </div>
        <div className="px-5 py-4 space-y-3">
          <Field label="Name *">
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              disabled={saving}
              placeholder='e.g., "Deepak Maheshwari"'
              className={INPUT}
              autoFocus
            />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Email">
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                disabled={saving}
                placeholder="name@domain.com"
                className={INPUT}
              />
            </Field>
            <Field label="Phone">
              <input
                type="tel"
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                disabled={saving}
                placeholder="(555) 555-5555"
                className={INPUT}
              />
            </Field>
          </div>
          <Field label="Role">
            <select
              value={role}
              onChange={(e) => setRole(e.target.value as ContactRole | '')}
              disabled={saving}
              className={INPUT + ' bg-white'}
            >
              <option value="">— Select a role —</option>
              {CONTACT_ROLES.map((r) => (
                <option key={r} value={r}>{r}</option>
              ))}
            </select>
          </Field>
          <Field label={`Linked Owner Entities (${selectedOwnerIds.size} selected)`}>
            <input
              type="text"
              value={ownerSearch}
              onChange={(e) => setOwnerSearch(e.target.value)}
              placeholder="Search owner entities…"
              disabled={saving}
              className={INPUT + ' mb-1'}
            />
            <div className="border border-gray-300 rounded max-h-40 overflow-y-auto bg-white">
              {filteredOwners.length === 0 ? (
                <div className="px-3 py-2 text-[11px] text-gray-500 italic">
                  No owner entities match your search.
                </div>
              ) : (
                filteredOwners.map((o) => (
                  <label
                    key={o.id}
                    className="flex items-center gap-2 px-2 py-1.5 hover:bg-teal-50 cursor-pointer text-xs"
                  >
                    <input
                      type="checkbox"
                      checked={selectedOwnerIds.has(String(o.id))}
                      onChange={() => toggleOwner(String(o.id))}
                      disabled={saving}
                    />
                    <span className="flex-1 min-w-0 truncate">{o.fields.Title}</span>
                    {o.fields.OwnerType && (
                      <span className="text-[10px] text-gray-500 flex-shrink-0">{o.fields.OwnerType}</span>
                    )}
                  </label>
                ))
              )}
            </div>
            <p className="text-[11px] text-gray-500 mt-1">
              Check every Owner entity this contact represents. A contact linked to multiple owners shows up
              under each owner's "Waiting on this owner" filter.
            </p>
          </Field>
          <Field label="Notes">
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              disabled={saving}
              rows={2}
              className={INPUT + ' resize-none'}
            />
          </Field>

          {error && (
            <div className="bg-red-50 border border-red-200 rounded p-2 text-xs text-red-800">
              {error}
            </div>
          )}
        </div>
        <div className="px-5 py-3 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={saving}
            className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded-md"
          >
            Cancel
          </button>
          <button
            onClick={handleSave}
            disabled={saving || !title.trim()}
            className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
          >
            {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {saving ? 'Saving…' : 'Add Contact'}
          </button>
        </div>
      </div>
    </div>
  );
}

const INPUT = 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
        {label}
      </label>
      {children}
    </div>
  );
}
