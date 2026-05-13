import { useState, useEffect, useRef, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  useSharePointList,
  LIST_NAMES,
  type Property,
  type ComplianceDeadline,
  type Ownership,
} from '../../lib/sharepoint';
import { Icon } from '../ui/Icon';

interface SearchResult {
  group: 'Properties' | 'Deadlines' | 'Ownership';
  id: string;
  title: string;
  subtitle?: string;
  route: string;
}

const MAX_RESULTS_PER_GROUP = 5;

export function GlobalSearch() {
  const navigate = useNavigate();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  // Only fetch when search is opened (lazy)
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const deadlines = useSharePointList<ComplianceDeadline>(LIST_NAMES.ComplianceDeadlines, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });

  // Global keyboard: Cmd/Ctrl+K opens search
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
        e.preventDefault();
        setOpen(true);
        setTimeout(() => inputRef.current?.focus(), 0);
      }
      if (e.key === 'Escape' && open) {
        setOpen(false);
        setQuery('');
        inputRef.current?.blur();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [open]);

  // Click outside to close
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const results: SearchResult[] = useMemo(() => {
    if (!query.trim()) return [];
    const q = query.toLowerCase().trim();
    const out: SearchResult[] = [];

    // Properties — match Title, PropertyAddress, LegalEntity
    if (properties.data) {
      const matched = properties.data
        .filter((p) => {
          const f = p.fields;
          const hay = `${f.Title ?? ''} ${f.PropertyAddress ?? ''} ${f.LegalEntity ?? ''}`.toLowerCase();
          return hay.includes(q);
        })
        .slice(0, MAX_RESULTS_PER_GROUP);
      matched.forEach((p) => {
        out.push({
          group: 'Properties',
          id: p.id,
          title: p.fields.Title ?? '(unnamed)',
          subtitle: p.fields.PropertyAddress || p.fields.LegalEntity || p.fields.cahpState,
          route: `/properties/${p.id}`,
        });
      });
    }

    // Deadlines — match Title, DeadlineType
    if (deadlines.data) {
      const matched = deadlines.data
        .filter((d) => {
          const f = d.fields;
          const hay = `${f.Title ?? ''} ${f.DeadlineType ?? ''}`.toLowerCase();
          return hay.includes(q);
        })
        .slice(0, MAX_RESULTS_PER_GROUP);
      matched.forEach((d) => {
        out.push({
          group: 'Deadlines',
          id: d.id,
          title: d.fields.Title ?? '(unnamed)',
          subtitle: `${d.fields.DeadlineType ?? '—'} · ${d.fields.DeadlineStatus ?? '—'}`,
          route: `/compliance/${d.id}`,
        });
      });
    }

    // Ownership — match Title, ParentEntity
    if (ownership.data) {
      const matched = ownership.data
        .filter((o) => {
          const f = o.fields;
          const hay = `${f.Title ?? ''} ${f.ParentEntity ?? ''}`.toLowerCase();
          return hay.includes(q);
        })
        .slice(0, MAX_RESULTS_PER_GROUP);
      matched.forEach((o) => {
        out.push({
          group: 'Ownership',
          id: o.id,
          title: o.fields.Title ?? '(unnamed)',
          subtitle: o.fields.RelationshipType ?? '—',
          route: `/ownership/${o.id}`,
        });
      });
    }

    return out;
  }, [query, properties.data, deadlines.data, ownership.data]);

  const grouped = useMemo(() => {
    const g: Record<string, SearchResult[]> = { Properties: [], Deadlines: [], Ownership: [] };
    results.forEach((r) => g[r.group].push(r));
    return g;
  }, [results]);

  const handleSelect = (route: string) => {
    setOpen(false);
    setQuery('');
    navigate(route);
  };

  return (
    <div ref={containerRef} className="flex-1 px-4 hidden md:flex justify-center relative">
      {!open ? (
        <button
          onClick={() => {
            setOpen(true);
            setTimeout(() => inputRef.current?.focus(), 0);
          }}
          className="w-full max-w-md flex items-center gap-2 px-3 py-1.5 rounded-md bg-white/10 hover:bg-white/15 text-teal-100 text-sm text-left transition-colors"
        >
          <Icon name="search" size={16} />
          <span className="flex-1">Search properties, deadlines, owners…</span>
          <span className="text-[11px] font-mono-data opacity-60">⌘K</span>
        </button>
      ) : (
        <div className="w-full max-w-md relative">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-md bg-white text-gray-900 text-sm">
            <Icon name="search" size={16} className="text-gray-400" />
            <input
              ref={inputRef}
              type="text"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search properties, deadlines, owners…"
              className="flex-1 outline-none bg-transparent"
            />
            <button
              onClick={() => {
                setOpen(false);
                setQuery('');
              }}
              className="text-[11px] font-mono-data text-gray-400 hover:text-gray-600"
            >
              Esc
            </button>
          </div>

          {query.trim() && (
            <div className="absolute top-full left-0 right-0 mt-1 bg-white text-gray-900 rounded-md shadow-xl border border-gray-200 max-h-[60vh] overflow-y-auto z-50">
              {results.length === 0 ? (
                <div className="px-4 py-6 text-center text-sm text-gray-500">
                  No results for "{query}"
                </div>
              ) : (
                <>
                  {(['Properties', 'Deadlines', 'Ownership'] as const).map((group) =>
                    grouped[group].length > 0 ? (
                      <div key={group}>
                        <div className="px-3 py-1.5 bg-gray-50 border-b border-gray-100 text-[11px] font-semibold text-gray-500 uppercase tracking-wider">
                          {group} ({grouped[group].length})
                        </div>
                        {grouped[group].map((r) => (
                          <button
                            key={r.id}
                            onClick={() => handleSelect(r.route)}
                            className="w-full text-left px-3 py-2 hover:bg-teal-50 border-b border-gray-100 last:border-b-0 transition-colors"
                          >
                            <div className="text-sm font-medium text-gray-900 truncate">{r.title}</div>
                            {r.subtitle && (
                              <div className="text-xs text-gray-500 truncate">{r.subtitle}</div>
                            )}
                          </button>
                        ))}
                      </div>
                    ) : null
                  )}
                </>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
