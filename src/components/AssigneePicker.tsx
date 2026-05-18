import { useId } from 'react';
import { TEAM_MEMBERS } from '../lib/roleMap';

/**
 * Assignee picker — an input with autocomplete suggestions drawn from the team
 * member roster, but accepting free text (vendors, owners, external counsel, etc.).
 *
 * Stores the display name as a string. Doesn't enforce that the value matches
 * a known team member — it's a hint, not a constraint.
 */
export function AssigneePicker({
  value,
  onChange,
  disabled,
  placeholder = 'Who is responsible? (start typing or pick from list)',
  className = '',
}: {
  value: string | undefined;
  onChange: (v: string) => void;
  disabled?: boolean;
  placeholder?: string;
  className?: string;
}) {
  const id = useId();
  return (
    <>
      <input
        type="text"
        list={id}
        value={value ?? ''}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
        placeholder={placeholder}
        className={className || 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500'}
      />
      <datalist id={id}>
        {TEAM_MEMBERS.map((m) => (
          <option key={m.email} value={m.name}>
            {m.email}
          </option>
        ))}
        <option value="Owner" />
        <option value="DOR" />
        <option value="Vendor" />
      </datalist>
    </>
  );
}
