import { useState, useMemo, useRef, useEffect } from 'react';
import { ALL_COUNTIES, parseCounties, joinCounties } from '../lib/counties';
import { Icon } from './ui/Icon';

/**
 * County multi-select for Property records.
 *
 * Single-family homes occasionally straddle two counties — and most rural
 * properties want to track both the parcel's home county and any adjacent
 * jurisdiction that touches it. So `cahpCounty` is a comma-joined string
 * (parsed here on read, joined on write).
 *
 * Display: chips for selected counties, click to remove.
 * Edit: type-ahead input that opens a searchable checkbox dropdown of all
 * SC + NC counties (~146 + "Other"). Closes on outside-click.
 */
export interface CountyMultiSelectProps {
  /** Stored comma-joined string. */
  value: string | undefined;
  /** Receives the new comma-joined string (empty when no counties selected). */
  onChange: (next: string) => void;
  disabled?: boolean;
}

export function CountyMultiSelect({ value, onChange, disabled }: CountyMultiSelectProps) {
  const selected = useMemo(() => parseCounties(value), [value]);
  const selectedSet = useMemo(() => new Set(selected), [selected]);
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Close dropdown on outside click
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

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return ALL_COUNTIES;
    return ALL_COUNTIES.filter((c) => c.toLowerCase().includes(q));
  }, [query]);

  const toggle = (county: string) => {
    const next = new Set(selectedSet);
    if (next.has(county)) next.delete(county);
    else next.add(county);
    onChange(joinCounties(Array.from(next)));
  };

  const remove = (county: string) => {
    const next = selected.filter((c) => c !== county);
    onChange(joinCounties(next));
  };

  return (
    <div ref={containerRef} className="relative">
      <div
        className={`w-full min-h-[34px] px-2 py-1 border border-gray-300 rounded bg-white flex flex-wrap items-center gap-1 cursor-text ${disabled ? 'opacity-60' : ''}`}
        onClick={() => !disabled && setOpen(true)}
      >
        {selected.length === 0 && (
          <span className="text-sm text-gray-400">— pick one or more counties —</span>
        )}
        {selected.map((c) => (
          <span
            key={c}
            className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-teal-50 text-teal-800 text-[11px] font-medium"
          >
            {c}
            {!disabled && (
              <button
                type="button"
                onClick={(e) => { e.stopPropagation(); remove(c); }}
                className="text-teal-700 hover:text-red-700"
                title={`Remove ${c}`}
              >
                <Icon name="x" size={10} />
              </button>
            )}
          </span>
        ))}
        {!disabled && (
          <input
            type="text"
            value={query}
            onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
            onFocus={() => setOpen(true)}
            placeholder={selected.length === 0 ? '' : 'Add another…'}
            className="flex-1 min-w-[80px] outline-none text-sm border-0 bg-transparent"
          />
        )}
      </div>

      {open && !disabled && (
        <div className="absolute z-30 mt-1 left-0 right-0 bg-white border border-gray-200 rounded-md shadow-lg max-h-60 overflow-y-auto">
          {filtered.length === 0 ? (
            <div className="px-3 py-2 text-xs text-gray-500">No counties match "{query}".</div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {filtered.slice(0, 60).map((c) => {
                const isOn = selectedSet.has(c);
                return (
                  <li key={c}>
                    <label className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs">
                      <input
                        type="checkbox"
                        checked={isOn}
                        onChange={() => toggle(c)}
                      />
                      <span className="flex-1">{c}</span>
                    </label>
                  </li>
                );
              })}
              {filtered.length > 60 && (
                <li className="px-3 py-1.5 text-[11px] text-gray-500 italic">
                  Showing first 60 of {filtered.length} — keep typing to narrow.
                </li>
              )}
            </ul>
          )}
        </div>
      )}
    </div>
  );
}
