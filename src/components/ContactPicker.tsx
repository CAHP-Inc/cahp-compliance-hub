import { useState, useMemo, useRef, useEffect } from 'react';
import {
  useSharePointList,
  createListItem,
  LIST_NAMES,
  type Contact,
  type ContactFields,
  type ContactRole,
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
              {selected.fields.ContactOwnerLookupId && (
                <span className="ml-1">
                  · {ownersById.get(String(selected.fields.ContactOwnerLookupId))?.fields.Title ?? 'unknown owner'}
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
              {filtered.slice(0, 20).map((c) => (
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
                      {c.fields.ContactOwnerLookupId && (
                        <span className="ml-1">
                          · {ownersById.get(String(c.fields.ContactOwnerLookupId))?.fields.Title ?? 'unknown owner'}
                        </span>
                      )}
                    </div>
                  </button>
                </li>
              ))}
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
  const [ownerLookupId, setOwnerLookupId] = useState(defaultOwnerId ?? '');
  const [notes, setNotes] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const sortedOwners = useMemo(
    () => [...(owners.data ?? [])].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    ),
    [owners.data],
  );

  const handleSave = async () => {
    if (!title.trim()) {
      setError('Contact name is required.');
      return;
    }
    setSaving(true);
    setError(null);
    try {
      const payload: ContactFields & { Title: string } = {
        Title: title.trim(),
        ContactEmail: email.trim() || undefined,
        ContactPhone: phone.trim() || undefined,
        ContactRole: role || undefined,
        ContactOwnerLookupId: ownerLookupId ? ownerLookupId : undefined,
        ContactNotes: notes.trim() || undefined,
      };
      const created = await createListItem<Contact>(LIST_NAMES.Contacts, payload as unknown as Record<string, unknown>);
      onSaved(String(created.id));
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
          <Field label="Linked Owner Entity">
            <select
              value={ownerLookupId}
              onChange={(e) => setOwnerLookupId(e.target.value)}
              disabled={saving}
              className={INPUT + ' bg-white'}
            >
              <option value="">— None / external —</option>
              {sortedOwners.map((o) => (
                <option key={o.id} value={String(o.id)}>{o.fields.Title}</option>
              ))}
            </select>
            <p className="text-[11px] text-gray-500 mt-1">
              If this contact represents an Owner entity in our system, link them here. Tasks waiting on this contact
              will surface on that Owner's detail page.
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
