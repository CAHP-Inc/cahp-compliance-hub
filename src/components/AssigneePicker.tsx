import { useEffect, useMemo, useRef, useState } from 'react';
import { TEAM_MEMBERS } from '../lib/roleMap';
import { useSharePointList, LIST_NAMES, type Contact } from '../lib/sharepoint';

/**
 * Assignee picker — text input + custom dropdown drawn from:
 *   1. TEAM_MEMBERS (signed-in app users)
 *   2. Contacts list (external people we ping — property owners, attorneys, vendors)
 *
 * Previously used <datalist>, which Chromium renders inconsistently —
 * some users had to click the arrow several times before all options
 * showed up, and Contacts were often hidden behind the team list. This
 * version manages its own dropdown so every option is always visible.
 *
 * Stores a free-text string (doesn't enforce a known-person value), so
 * legacy items and ad-hoc assignees still work.
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
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });

  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Close when the user clicks outside the picker
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

  // Build the combined option list. Team first, contacts after, deduped by
  // name so a person with both a TEAM_MEMBERS entry and a Contact row only
  // appears once (team wins for the label).
  const allOptions = useMemo(() => {
    const seen = new Set<string>();
    const out: { name: string; sub: string; kind: 'team' | 'contact' }[] = [];
    for (const m of TEAM_MEMBERS) {
      const key = m.name.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ name: m.name, sub: m.email, kind: 'team' });
    }
    for (const c of contacts.data ?? []) {
      const name = (c.fields.Title ?? '').trim();
      if (!name) continue;
      const key = name.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      const subParts: string[] = [];
      if (c.fields.ContactEmail) subParts.push(c.fields.ContactEmail);
      if (c.fields.ContactRole) subParts.push(c.fields.ContactRole);
      out.push({ name, sub: subParts.join(' · ') || 'contact', kind: 'contact' });
    }
    return out;
  }, [contacts.data]);

  // Live filter against whatever the user has typed
  const filtered = useMemo(() => {
    const q = (value ?? '').trim().toLowerCase();
    if (!q) return allOptions;
    return allOptions.filter(
      (o) => o.name.toLowerCase().includes(q) || o.sub.toLowerCase().includes(q),
    );
  }, [allOptions, value]);

  const teamOptions = filtered.filter((o) => o.kind === 'team');
  const contactOptions = filtered.filter((o) => o.kind === 'contact');

  return (
    <div ref={containerRef} className="relative w-full">
      <input
        type="text"
        value={value ?? ''}
        onChange={(e) => {
          onChange(e.target.value);
          if (!open) setOpen(true);
        }}
        onFocus={() => setOpen(true)}
        onBlur={onBlur}
        onKeyDown={(e) => {
          if (e.key === 'Escape') setOpen(false);
          onKeyDown?.(e);
        }}
        disabled={disabled}
        placeholder={placeholder}
        className={className || 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500'}
      />
      {!disabled && open && (
        <div
          className="absolute z-50 mt-1 left-0 right-0 bg-white border border-gray-200 rounded-md shadow-lg max-h-64 overflow-y-auto"
          // Use onMouseDown so the click registers before the input's onBlur fires
          // and tears down state — otherwise the picker closes before onChange runs.
          onMouseDown={(e) => e.preventDefault()}
        >
          {filtered.length === 0 ? (
            <div className="px-3 py-2 text-xs text-gray-500 italic">
              No matches. Press Enter to use "{value}" as a free-text assignee.
            </div>
          ) : (
            <>
              {teamOptions.length > 0 && (
                <Section label="Team" options={teamOptions} onPick={(n) => { onChange(n); setOpen(false); }} />
              )}
              {contactOptions.length > 0 && (
                <Section label="Contacts" options={contactOptions} onPick={(n) => { onChange(n); setOpen(false); }} />
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
}

function Section({
  label,
  options,
  onPick,
}: {
  label: string;
  options: { name: string; sub: string }[];
  onPick: (name: string) => void;
}) {
  return (
    <div>
      <div className="px-3 py-1 bg-gray-50 text-[10px] font-semibold text-gray-500 uppercase tracking-wider sticky top-0">
        {label} ({options.length})
      </div>
      <ul className="divide-y divide-gray-100">
        {options.map((o) => (
          <li key={`${label}-${o.name}`}>
            <button
              type="button"
              onClick={() => onPick(o.name)}
              className="w-full text-left px-3 py-1.5 hover:bg-teal-50 cursor-pointer"
            >
              <div className="text-sm text-gray-900">{o.name}</div>
              {o.sub && <div className="text-[11px] text-gray-500 font-mono-data truncate">{o.sub}</div>}
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
