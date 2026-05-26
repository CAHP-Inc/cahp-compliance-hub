import { useId } from 'react';
import { TEAM_MEMBERS } from '../lib/roleMap';
import { useSharePointList, LIST_NAMES, type Contact } from '../lib/sharepoint';

/**
 * Assignee picker — a text input with autocomplete drawn from:
 *   1. Hardcoded TEAM_MEMBERS (signed-in app users)
 *   2. Contacts list (external people we ping — property owners, attorneys, vendors)
 *
 * Stores a free-text string; doesn't enforce that the value matches a known
 * person. That keeps backward compatibility with legacy items and supports
 * ad-hoc assignees (vendors with no Contact record yet, etc.).
 *
 * Convention: the dropdown shows the display name as the option value and
 * the email as the visible label, so picking an option fills the input with
 * the name. This keeps "Waiting on this owner" filters working — the
 * OwnerDetail page matches AssignedTo against any linked Contact's
 * name OR email, case-insensitive.
 */
export function AssigneePicker({
  value,
  onChange,
  onBlur,
  onKeyDown,
  disabled,
  placeholder = 'Who is responsible? (start typing or pick from list)',
  className = '',
}: {
  value: string | undefined;
  onChange: (v: string) => void;
  onBlur?: (e: React.FocusEvent<HTMLInputElement>) => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLInputElement>) => void;
  disabled?: boolean;
  placeholder?: string;
  className?: string;
}) {
  const id = useId();
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });

  return (
    <>
      <input
        type="text"
        list={id}
        value={value ?? ''}
        onChange={(e) => onChange(e.target.value)}
        onBlur={onBlur}
        onKeyDown={onKeyDown}
        disabled={disabled}
        placeholder={placeholder}
        className={className || 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500'}
      />
      <datalist id={id}>
        {TEAM_MEMBERS.map((m) => (
          <option key={`team-${m.email}`} value={m.name}>
            {m.email} (team)
          </option>
        ))}
        {(contacts.data ?? []).map((c) => {
          const name = c.fields.Title ?? '';
          if (!name) return null;
          const labelParts: string[] = [];
          if (c.fields.ContactEmail) labelParts.push(c.fields.ContactEmail);
          if (c.fields.ContactRole) labelParts.push(c.fields.ContactRole);
          return (
            <option key={`contact-${c.id}`} value={name}>
              {labelParts.join(' · ') || 'contact'}
            </option>
          );
        })}
      </datalist>
    </>
  );
}
