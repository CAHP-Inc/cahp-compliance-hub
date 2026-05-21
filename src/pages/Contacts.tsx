import { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import {
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Contact,
  type ContactFields,
  type ContactRole,
  type ContactOwnerLink,
  type Owner,
  type OutstandingItem,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { NewContactModal } from '../components/ContactPicker';

/**
 * Contacts directory — people we communicate with about properties.
 *
 * Distinct from the Owners list (entity records). A Contact CAN be linked to
 * an Owner via ContactOwnerLookupId; that linkage is what powers the
 * "waiting on this owner" filter on OwnerDetail.
 */

const CONTACT_ROLES: ContactRole[] = [
  'Property Owner',
  'Sponsor',
  'Attorney',
  'Accountant',
  'Property Manager',
  'Vendor',
  'Lender',
  'DOR',
  'Other',
];

const ROLE_STYLES: Record<ContactRole, string> = {
  'Property Owner':   'bg-blue-100 text-blue-800',
  Sponsor:            'bg-amber-100 text-amber-800',
  Attorney:           'bg-indigo-100 text-indigo-800',
  Accountant:         'bg-emerald-100 text-emerald-800',
  'Property Manager': 'bg-teal-100 text-teal-800',
  Vendor:             'bg-gray-100 text-gray-700',
  Lender:             'bg-rose-100 text-rose-800',
  DOR:                'bg-purple-100 text-purple-800',
  Other:              'bg-gray-100 text-gray-700',
};

export function Contacts() {
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownerLinks = useSharePointList<ContactOwnerLink>(LIST_NAMES.ContactOwnerLinks, { top: 2000 });
  const items = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 500 });

  const [search, setSearch] = useState('');
  const [roleFilter, setRoleFilter] = useState<ContactRole | 'All'>('All');
  const [creating, setCreating] = useState(false);
  const [editing, setEditing] = useState<Contact | null>(null);

  const ownersById = useMemo(() => {
    const m = new Map<string, Owner>();
    (owners.data ?? []).forEach((o) => m.set(String(o.id), o));
    return m;
  }, [owners.data]);

  // contactId → list of {ownerId, ownerTitle} pairs (junction + legacy single field)
  const linkedOwnersByContact = useMemo(() => {
    const map = new Map<string, { id: string; title: string }[]>();
    const add = (cId: string, oId: string) => {
      const title = ownersById.get(oId)?.fields.Title;
      if (!title) return;
      if (!map.has(cId)) map.set(cId, []);
      const list = map.get(cId)!;
      if (!list.some((e) => e.id === oId)) list.push({ id: oId, title });
    };
    (ownerLinks.data ?? []).forEach((l) => {
      if (l.fields.ContactLookupId && l.fields.OwnerLookupId) {
        add(String(l.fields.ContactLookupId), String(l.fields.OwnerLookupId));
      }
    });
    (contacts.data ?? []).forEach((c) => {
      if (c.fields.ContactOwnerLookupId) {
        add(String(c.id), String(c.fields.ContactOwnerLookupId));
      }
    });
    return map;
  }, [ownerLinks.data, contacts.data, ownersById]);

  // Pre-count open items per contact (by name OR email match against AssignedTo)
  const openItemCountByContact = useMemo(() => {
    const isClosed = (s: string | undefined) =>
      s === 'Done' || s === 'Received' || s === 'Not Applicable';
    const open = (items.data ?? []).filter((i) => !isClosed(i.fields.ItemStatus));
    const counts = new Map<string, number>();
    for (const c of contacts.data ?? []) {
      const name = (c.fields.Title ?? '').trim().toLowerCase();
      const email = (c.fields.ContactEmail ?? '').trim().toLowerCase();
      if (!name && !email) continue;
      const count = open.filter((i) => {
        const a = (i.fields.AssignedTo ?? '').trim().toLowerCase();
        if (!a) return false;
        return (!!name && a === name) || (!!email && a === email);
      }).length;
      counts.set(String(c.id), count);
    }
    return counts;
  }, [contacts.data, items.data]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    let list = [...(contacts.data ?? [])];
    if (q) {
      list = list.filter((c) => {
        const hay = `${c.fields.Title ?? ''} ${c.fields.ContactEmail ?? ''} ${c.fields.ContactPhone ?? ''} ${c.fields.ContactNotes ?? ''}`.toLowerCase();
        return hay.includes(q);
      });
    }
    if (roleFilter !== 'All') {
      list = list.filter((c) => c.fields.ContactRole === roleFilter);
    }
    return list.sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [contacts.data, search, roleFilter]);

  const loading = contacts.loading || owners.loading;
  const error = contacts.error;

  return (
    <div>
      <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Contacts</h1>
          <p className="text-sm text-gray-500 mt-1">
            People we communicate with about properties — owners, attorneys, vendors. Used to populate the assignee picker on tasks
            and to surface what's waiting on each person.
          </p>
        </div>
        <button
          onClick={() => setCreating(true)}
          className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
        >
          <Icon name="plus" size={14} />
          New Contact
        </button>
      </div>

      <div className="bg-white border border-gray-200 rounded-lg shadow-card mb-6 p-3 flex items-center gap-3 flex-wrap">
        <input
          type="text"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search by name, email, phone, or notes…"
          className="flex-1 min-w-[200px] px-3 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500"
        />
        <select
          value={roleFilter}
          onChange={(e) => setRoleFilter(e.target.value as ContactRole | 'All')}
          className="px-2 py-1.5 border border-gray-300 rounded text-sm bg-white focus:outline-none focus:border-teal-500"
        >
          <option value="All">All roles</option>
          {CONTACT_ROLES.map((r) => (
            <option key={r} value={r}>{r}</option>
          ))}
        </select>
        <div className="text-xs text-gray-500">
          {filtered.length} of {contacts.data?.length ?? 0} contacts
        </div>
      </div>

      {loading ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center text-sm text-gray-500 shadow-card">
          Loading contacts…
        </div>
      ) : error ? (
        <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-sm text-red-900">
          {String(error)}
          <div className="mt-2 text-xs text-red-700">
            If the Contacts list doesn't exist in SharePoint yet, see the provisioning notes in <code>CONTACTS_PROVISIONING.md</code>.
          </div>
        </div>
      ) : filtered.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500 mb-2">
            {search || roleFilter !== 'All' ? 'No contacts match your filter.' : 'No contacts in the system yet.'}
          </p>
          {!search && roleFilter === 'All' && (
            <button
              onClick={() => setCreating(true)}
              className="text-xs text-teal-700 hover:text-teal-900 font-medium"
            >
              Add your first contact →
            </button>
          )}
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Name</th>
                <th className="px-4 py-3 text-left">Email</th>
                <th className="px-4 py-3 text-left">Phone</th>
                <th className="px-4 py-3 text-left">Role</th>
                <th className="px-4 py-3 text-left">Linked Owner</th>
                <th className="px-4 py-3 text-right">Open Items</th>
                <th className="px-4 py-3 text-right w-20"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((c) => {
                const linkedOwners = linkedOwnersByContact.get(String(c.id)) ?? [];
                const openCount = openItemCountByContact.get(String(c.id)) ?? 0;
                return (
                  <tr key={c.id} className="hover:bg-gray-50">
                    <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                    <td className="px-4 py-3 text-xs text-gray-700 font-mono-data">
                      {c.fields.ContactEmail ? (
                        <a href={`mailto:${c.fields.ContactEmail}`} className="text-teal-700 hover:text-teal-900">
                          {c.fields.ContactEmail}
                        </a>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3 text-xs text-gray-700 font-mono-data">
                      {c.fields.ContactPhone || '—'}
                    </td>
                    <td className="px-4 py-3">
                      {c.fields.ContactRole && (
                        <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${ROLE_STYLES[c.fields.ContactRole]}`}>
                          {c.fields.ContactRole}
                        </span>
                      )}
                    </td>
                    <td className="px-4 py-3 text-xs">
                      {linkedOwners.length === 0 ? (
                        <span className="text-gray-400 italic">—</span>
                      ) : (
                        <div className="flex flex-wrap gap-1">
                          {linkedOwners.map((o) => (
                            <Link
                              key={o.id}
                              to={`/owners/${o.id}`}
                              className="px-1.5 py-0.5 rounded bg-teal-50 text-teal-800 hover:bg-teal-100 text-[11px] font-medium"
                            >
                              {o.title}
                            </Link>
                          ))}
                        </div>
                      )}
                    </td>
                    <td className="px-4 py-3 text-right">
                      {openCount > 0 ? (
                        <span className="bg-amber-100 text-amber-900 font-mono-data text-xs font-semibold px-2 py-0.5 rounded">
                          {openCount}
                        </span>
                      ) : (
                        <span className="text-gray-300">—</span>
                      )}
                    </td>
                    <td className="px-4 py-3 text-right">
                      <button
                        onClick={() => setEditing(c)}
                        className="text-[11px] text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50"
                      >
                        Edit
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {creating && (
        <NewContactModal
          onClose={() => setCreating(false)}
          onSaved={() => {
            contacts.refetch?.();
            setCreating(false);
          }}
        />
      )}

      {editing && (
        <EditContactModal
          contact={editing}
          initialOwnerIds={
            (linkedOwnersByContact.get(String(editing.id)) ?? []).map((o) => o.id)
          }
          existingLinkRows={ownerLinks.data ?? []}
          onClose={() => setEditing(null)}
          onSaved={() => {
            contacts.refetch?.();
            ownerLinks.refetch?.();
            setEditing(null);
          }}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Edit Contact modal — separate from New so we can show delete + linkage
// ---------------------------------------------------------------------------

function EditContactModal({
  contact,
  initialOwnerIds,
  existingLinkRows,
  onClose,
  onSaved,
}: {
  contact: Contact;
  initialOwnerIds: string[];
  existingLinkRows: ContactOwnerLink[];
  onClose: () => void;
  onSaved: () => void;
}) {
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const [title, setTitle] = useState(contact.fields.Title ?? '');
  const [email, setEmail] = useState(contact.fields.ContactEmail ?? '');
  const [phone, setPhone] = useState(contact.fields.ContactPhone ?? '');
  const [role, setRole] = useState<ContactRole | ''>(contact.fields.ContactRole ?? '');
  const [selectedOwnerIds, setSelectedOwnerIds] = useState<Set<string>>(
    new Set(initialOwnerIds),
  );
  const [ownerSearch, setOwnerSearch] = useState('');
  const [notes, setNotes] = useState(contact.fields.ContactNotes ?? '');
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
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
      // Treat the first selected owner as the legacy "primary" (keeps the
      // single ContactOwner SP column meaningful for default views)
      const primaryOwnerId = selectedOwnerIds.size > 0
        ? Array.from(selectedOwnerIds)[0]
        : null;

      const payload: Partial<ContactFields> = {
        Title: title.trim(),
        ContactEmail: email.trim() || null as unknown as undefined,
        ContactPhone: phone.trim() || null as unknown as undefined,
        ContactRole: (role || null) as ContactRole | undefined,
        ContactOwnerLookupId: (primaryOwnerId ?? null) as unknown as undefined,
        ContactNotes: notes.trim() || (null as unknown as undefined),
      };
      await updateListItem(LIST_NAMES.Contacts, contact.id, payload as Record<string, unknown>);

      // Diff junction rows: add new, remove unselected
      const initialSet = new Set(initialOwnerIds);
      const toAdd = Array.from(selectedOwnerIds).filter((id) => !initialSet.has(id));
      const toRemove = Array.from(initialSet).filter((id) => !selectedOwnerIds.has(id));

      const myJunctionRows = existingLinkRows.filter(
        (l) => String(l.fields.ContactLookupId ?? '') === String(contact.id),
      );

      for (const ownerId of toAdd) {
        await createListItem(LIST_NAMES.ContactOwnerLinks, {
          Title: `Contact ${contact.id} ↔ Owner ${ownerId}`,
          ContactLookupId: Number(contact.id),
          OwnerLookupId: Number(ownerId),
        });
      }
      for (const ownerId of toRemove) {
        const row = myJunctionRows.find(
          (l) => String(l.fields.OwnerLookupId ?? '') === String(ownerId),
        );
        if (row) {
          await deleteListItem(LIST_NAMES.ContactOwnerLinks, row.id);
        }
      }

      onSaved();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Delete contact "${contact.fields.Title}"? This does NOT change anything on outstanding items assigned to them — those keep the assignee text as-is.`)) return;
    setDeleting(true);
    try {
      // Clean up junction rows pointing at this contact first
      const myJunctionRows = existingLinkRows.filter(
        (l) => String(l.fields.ContactLookupId ?? '') === String(contact.id),
      );
      for (const row of myJunctionRows) {
        await deleteListItem(LIST_NAMES.ContactOwnerLinks, row.id);
      }
      await deleteListItem(LIST_NAMES.Contacts, contact.id);
      onSaved();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setDeleting(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !saving && !deleting) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-md w-full max-h-[90vh] overflow-y-auto">
        <div className="px-5 py-3 border-b border-gray-200">
          <h2 className="text-base font-bold text-teal-700">Edit Contact</h2>
        </div>
        <div className="px-5 py-4 space-y-3">
          <Field label="Name *">
            <input type="text" value={title} onChange={(e) => setTitle(e.target.value)} disabled={saving || deleting} className={EDIT_INPUT} />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Email">
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} disabled={saving || deleting} className={EDIT_INPUT} />
            </Field>
            <Field label="Phone">
              <input type="tel" value={phone} onChange={(e) => setPhone(e.target.value)} disabled={saving || deleting} className={EDIT_INPUT} />
            </Field>
          </div>
          <Field label="Role">
            <select value={role} onChange={(e) => setRole(e.target.value as ContactRole | '')} disabled={saving || deleting} className={EDIT_INPUT + ' bg-white'}>
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
              disabled={saving || deleting}
              className={EDIT_INPUT + ' mb-1'}
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
                      disabled={saving || deleting}
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
              Check every Owner entity this contact represents.
            </p>
          </Field>
          <Field label="Notes">
            <textarea value={notes} onChange={(e) => setNotes(e.target.value)} disabled={saving || deleting} rows={2} className={EDIT_INPUT + ' resize-none'} />
          </Field>

          {error && (
            <div className="bg-red-50 border border-red-200 rounded p-2 text-xs text-red-800">{error}</div>
          )}
        </div>
        <div className="px-5 py-3 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-between gap-2">
          <button onClick={handleDelete} disabled={saving || deleting} className="text-xs text-error hover:text-red-700 font-medium px-3 py-1.5 rounded hover:bg-red-50 disabled:opacity-50">
            {deleting ? 'Deleting…' : 'Delete'}
          </button>
          <div className="flex items-center gap-2">
            <button onClick={onClose} disabled={saving || deleting} className="px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-200 rounded-md">Cancel</button>
            <button onClick={handleSave} disabled={saving || deleting || !title.trim()} className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white px-4 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5">
              {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
              {saving ? 'Saving…' : 'Save'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

const EDIT_INPUT = 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">{label}</label>
      {children}
    </div>
  );
}
