import { useEffect, useMemo, useState } from 'react';
import type {
  ReportDescriptor,
  ReportFormat,
  ReportScope,
  ReportScopeKind,
  RunOptions,
} from '../lib/reports';
import type { CahpState, CahpTaxYear, Owner, Property } from '../lib/sharepoint';
import { Icon } from './ui/Icon';
import { toDateInputValue } from '../lib/dates';

/**
 * Run Report config modal.
 *
 * Renders only the controls the descriptor's params block declares — scope
 * picker, date range, tax year, expiration window, quarter, format, and
 * internal-columns toggle. The submit handler hands the assembled RunOptions
 * back to the page, which dispatches to the runner.
 */
export interface RunReportModalProps {
  descriptor: ReportDescriptor;
  properties: Property[];
  owners: Owner[];
  onClose: () => void;
  onRun: (opts: RunOptions) => Promise<void> | void;
}

const FORMAT_LABEL: Record<ReportFormat, string> = {
  csv: 'CSV',
  xlsx: 'Excel (.xlsx)',
  json: 'JSON',
  pdf: 'PDF',
};

const SCOPE_LABEL: Record<ReportScopeKind, string> = {
  portfolio: 'Entire portfolio',
  property: 'A single property',
  owner: 'A single owner',
  state: 'One state',
};

const TAX_YEARS: CahpTaxYear[] = ['2023', '2024', '2025', '2026', '2027', '2028'];
const STATES: CahpState[] = ['SC', 'NC'];

const todayISO = () => toDateInputValue(new Date());

export function RunReportModal({
  descriptor,
  properties,
  owners,
  onClose,
  onRun,
}: RunReportModalProps) {
  const allowedScopes: ReportScopeKind[] = useMemo(() => {
    return descriptor.params.scope ?? ['portfolio'];
  }, [descriptor.params.scope]);

  const [scopeKind, setScopeKind] = useState<ReportScopeKind>(allowedScopes[0]);
  const [propertyId, setPropertyId] = useState<string>('');
  const [ownerId, setOwnerId] = useState<string>('');
  const [stateCode, setStateCode] = useState<CahpState>('SC');
  const [propertySearch, setPropertySearch] = useState('');
  const [ownerSearch, setOwnerSearch] = useState('');

  // Date range: default to YTD if asked for
  const [dateFrom, setDateFrom] = useState<string>(() => {
    const d = new Date();
    return `${d.getFullYear()}-01-01`;
  });
  const [dateTo, setDateTo] = useState<string>(todayISO());

  // Tax year
  const [taxYear, setTaxYear] = useState<CahpTaxYear>(() => {
    const y = String(new Date().getFullYear()) as CahpTaxYear;
    return TAX_YEARS.includes(y) ? y : TAX_YEARS[TAX_YEARS.length - 2];
  });

  // Expiration window
  const [expirationWindow, setExpirationWindow] = useState<90 | 180 | 365>(180);

  // Quarter
  const now = new Date();
  const [quarterYear, setQuarterYear] = useState<number>(now.getFullYear());
  const [quarter, setQuarter] = useState<1 | 2 | 3 | 4>(
    (Math.floor(now.getMonth() / 3) + 1) as 1 | 2 | 3 | 4,
  );

  // Format
  const [format, setFormat] = useState<ReportFormat>(descriptor.defaultFormat);

  // Internal columns — default by audience (internal: on, owner: off)
  const [includeInternalColumns, setIncludeInternalColumns] = useState<boolean>(
    descriptor.audience === 'internal',
  );

  // Reset when descriptor changes (modal is keyed on descriptor.id but be safe)
  useEffect(() => {
    setScopeKind(allowedScopes[0]);
    setFormat(descriptor.defaultFormat);
    setIncludeInternalColumns(descriptor.audience === 'internal');
    setPropertyId('');
    setOwnerId('');
  }, [descriptor.id, allowedScopes, descriptor.defaultFormat, descriptor.audience]);

  const sortedProperties = useMemo(
    () => [...properties].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    ),
    [properties],
  );
  const filteredProperties = useMemo(() => {
    const q = propertySearch.trim().toLowerCase();
    if (!q) return sortedProperties;
    return sortedProperties.filter((p) => {
      const hay = `${p.fields.Title ?? ''} ${p.fields.cahpCounty ?? ''} ${p.fields.cahpState ?? ''}`.toLowerCase();
      return hay.includes(q);
    });
  }, [sortedProperties, propertySearch]);

  const sortedOwners = useMemo(
    () => [...owners].sort((a, b) =>
      (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''),
    ),
    [owners],
  );
  const filteredOwners = useMemo(() => {
    const q = ownerSearch.trim().toLowerCase();
    if (!q) return sortedOwners;
    return sortedOwners.filter((o) => {
      const hay = `${o.fields.Title ?? ''} ${o.fields.OwnerType ?? ''} ${o.fields.OwnerState ?? ''} ${o.fields.SponsorName ?? ''}`.toLowerCase();
      return hay.includes(q);
    });
  }, [sortedOwners, ownerSearch]);

  const [running, setRunning] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const canRun = useMemo(() => {
    if (running) return false;
    if (scopeKind === 'property' && !propertyId) return false;
    if (scopeKind === 'owner' && !ownerId) return false;
    if (descriptor.params.dateRange && (!dateFrom || !dateTo)) return false;
    return true;
  }, [running, scopeKind, propertyId, ownerId, descriptor.params.dateRange, dateFrom, dateTo]);

  const buildScope = (): ReportScope => {
    if (scopeKind === 'property') {
      const p = properties.find((x) => String(x.id) === propertyId);
      return { kind: 'property', propertyId, propertyTitle: p?.fields.Title };
    }
    if (scopeKind === 'owner') {
      const o = owners.find((x) => String(x.id) === ownerId);
      return { kind: 'owner', ownerId, ownerTitle: o?.fields.Title };
    }
    if (scopeKind === 'state') {
      return { kind: 'state', state: stateCode };
    }
    return { kind: 'portfolio' };
  };

  const handleRun = async () => {
    setError(null);
    setRunning(true);
    try {
      await onRun({
        format,
        scope: buildScope(),
        dateFrom: descriptor.params.dateRange ? dateFrom : undefined,
        dateTo: descriptor.params.dateRange ? dateTo : undefined,
        taxYear: descriptor.params.taxYear ? taxYear : undefined,
        expirationWindow: descriptor.params.expirationWindow ? expirationWindow : undefined,
        quarter: descriptor.params.quarter ? { year: quarterYear, q: quarter } : undefined,
        includeInternalColumns: descriptor.params.internalColumnsToggle
          ? includeInternalColumns
          : undefined,
      });
      // onRun resolves only after download/email handoff — close on success
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setRunning(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4"
      onClick={(e) => {
        if (e.target === e.currentTarget && !running) onClose();
      }}
    >
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[92vh] flex flex-col">
        <div className="px-6 py-4 border-b border-gray-200">
          <h2 className="text-lg font-bold text-teal-700">{descriptor.name}</h2>
          <p className="text-xs text-gray-600 mt-0.5">{descriptor.description}</p>
        </div>

        <div className="px-6 py-4 flex-1 overflow-y-auto space-y-4">
          {/* Scope */}
          {allowedScopes.length > 1 && (
            <FieldGroup label="Scope">
              <div className="flex flex-wrap gap-1.5">
                {allowedScopes.map((kind) => (
                  <button
                    key={kind}
                    type="button"
                    onClick={() => setScopeKind(kind)}
                    disabled={running}
                    className={`px-3 py-1.5 text-xs font-medium rounded-md border transition-colors ${
                      scopeKind === kind
                        ? 'bg-teal-700 border-teal-700 text-white'
                        : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
                    }`}
                  >
                    {SCOPE_LABEL[kind]}
                  </button>
                ))}
              </div>
            </FieldGroup>
          )}

          {scopeKind === 'property' && (
            <FieldGroup label="Pick a property" required>
              <input
                type="text"
                value={propertySearch}
                onChange={(e) => setPropertySearch(e.target.value)}
                placeholder="Search by name, county, or state…"
                disabled={running}
                className={inputClass + ' mb-1'}
              />
              <div className="border border-gray-300 rounded max-h-48 overflow-y-auto bg-white">
                {filteredProperties.length === 0 ? (
                  <div className="px-3 py-2 text-[11px] text-gray-500 italic">
                    No properties match.
                  </div>
                ) : (
                  filteredProperties.map((p) => (
                    <label
                      key={p.id}
                      className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs"
                    >
                      <input
                        type="radio"
                        name="property-pick"
                        checked={propertyId === String(p.id)}
                        onChange={() => setPropertyId(String(p.id))}
                        disabled={running}
                      />
                      <span className="flex-1 truncate">{p.fields.Title}</span>
                      {p.fields.cahpState && (
                        <span className="text-[10px] text-gray-500 flex-shrink-0">
                          {p.fields.cahpState}
                        </span>
                      )}
                    </label>
                  ))
                )}
              </div>
            </FieldGroup>
          )}

          {scopeKind === 'owner' && (
            <FieldGroup label="Pick an owner" required>
              <input
                type="text"
                value={ownerSearch}
                onChange={(e) => setOwnerSearch(e.target.value)}
                placeholder="Search by name, type, sponsor, or state…"
                disabled={running}
                className={inputClass + ' mb-1'}
              />
              <div className="border border-gray-300 rounded max-h-48 overflow-y-auto bg-white">
                {filteredOwners.length === 0 ? (
                  <div className="px-3 py-2 text-[11px] text-gray-500 italic">
                    No owners match.
                  </div>
                ) : (
                  filteredOwners.map((o) => (
                    <label
                      key={o.id}
                      className="flex items-center gap-2 px-2 py-1 hover:bg-teal-50 cursor-pointer text-xs"
                    >
                      <input
                        type="radio"
                        name="owner-pick"
                        checked={ownerId === String(o.id)}
                        onChange={() => setOwnerId(String(o.id))}
                        disabled={running}
                      />
                      <span className="flex-1 truncate">{o.fields.Title}</span>
                      {o.fields.OwnerType && (
                        <span className="text-[10px] text-gray-500 flex-shrink-0">
                          {o.fields.OwnerType}
                        </span>
                      )}
                    </label>
                  ))
                )}
              </div>
            </FieldGroup>
          )}

          {scopeKind === 'state' && (
            <FieldGroup label="Pick a state" required>
              <div className="flex gap-1.5">
                {STATES.map((s) => (
                  <button
                    key={s}
                    type="button"
                    onClick={() => setStateCode(s)}
                    disabled={running}
                    className={`px-3 py-1.5 text-xs font-medium rounded-md border ${
                      stateCode === s
                        ? 'bg-teal-700 border-teal-700 text-white'
                        : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
                    }`}
                  >
                    {s}
                  </button>
                ))}
              </div>
            </FieldGroup>
          )}

          {/* Date range */}
          {descriptor.params.dateRange && (
            <FieldGroup label="Date range">
              <div className="flex items-center gap-2">
                <input
                  type="date"
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                  disabled={running}
                  className={inputClass + ' flex-1'}
                />
                <span className="text-xs text-gray-500">to</span>
                <input
                  type="date"
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                  disabled={running}
                  className={inputClass + ' flex-1'}
                />
              </div>
            </FieldGroup>
          )}

          {/* Tax year */}
          {descriptor.params.taxYear && (
            <FieldGroup label="Tax year">
              <select
                value={taxYear}
                onChange={(e) => setTaxYear(e.target.value as CahpTaxYear)}
                disabled={running}
                className={inputClass + ' bg-white'}
              >
                {TAX_YEARS.map((y) => (
                  <option key={y} value={y}>
                    {y}
                  </option>
                ))}
              </select>
            </FieldGroup>
          )}

          {/* Expiration window */}
          {descriptor.params.expirationWindow && (
            <FieldGroup label="Window">
              <div className="flex gap-1.5">
                {([90, 180, 365] as const).map((w) => (
                  <button
                    key={w}
                    type="button"
                    onClick={() => setExpirationWindow(w)}
                    disabled={running}
                    className={`px-3 py-1.5 text-xs font-medium rounded-md border ${
                      expirationWindow === w
                        ? 'bg-teal-700 border-teal-700 text-white'
                        : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
                    }`}
                  >
                    Next {w} days
                  </button>
                ))}
              </div>
            </FieldGroup>
          )}

          {/* Quarter */}
          {descriptor.params.quarter && (
            <FieldGroup label="Statement period">
              <div className="flex items-center gap-2">
                <select
                  value={quarter}
                  onChange={(e) =>
                    setQuarter(Number(e.target.value) as 1 | 2 | 3 | 4)
                  }
                  disabled={running}
                  className={inputClass + ' bg-white w-24'}
                >
                  <option value={1}>Q1</option>
                  <option value={2}>Q2</option>
                  <option value={3}>Q3</option>
                  <option value={4}>Q4</option>
                </select>
                <input
                  type="number"
                  min={2020}
                  max={2099}
                  value={quarterYear}
                  onChange={(e) => setQuarterYear(parseInt(e.target.value, 10) || quarterYear)}
                  disabled={running}
                  className={inputClass + ' w-24'}
                />
              </div>
            </FieldGroup>
          )}

          {/* Format */}
          {descriptor.supportedFormats.length > 1 && (
            <FieldGroup label="Format">
              <div className="flex flex-wrap gap-1.5">
                {descriptor.supportedFormats.map((f) => (
                  <button
                    key={f}
                    type="button"
                    onClick={() => setFormat(f)}
                    disabled={running}
                    className={`px-3 py-1.5 text-xs font-medium rounded-md border ${
                      format === f
                        ? 'bg-teal-700 border-teal-700 text-white'
                        : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
                    }`}
                  >
                    {FORMAT_LABEL[f]}
                  </button>
                ))}
              </div>
            </FieldGroup>
          )}

          {/* Internal columns */}
          {descriptor.params.internalColumnsToggle && (
            <FieldGroup label="Detail level">
              <label className="inline-flex items-center gap-2 cursor-pointer text-xs">
                <input
                  type="checkbox"
                  checked={includeInternalColumns}
                  onChange={(e) => setIncludeInternalColumns(e.target.checked)}
                  disabled={running}
                />
                <span>
                  Include internal-only columns (notes, AssignedTo, internal status detail).
                  Uncheck before sending the file to an owner.
                </span>
              </label>
            </FieldGroup>
          )}

          {error && (
            <div className="bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">
              {error}
            </div>
          )}
        </div>

        <div className="px-6 py-4 border-t border-gray-200 bg-gray-50 rounded-b-lg flex items-center justify-end gap-2">
          <button
            onClick={onClose}
            disabled={running}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-100 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleRun}
            disabled={!canRun}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:cursor-not-allowed"
          >
            {running && (
              <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
            )}
            <Icon name="download" size={13} />
            {running ? 'Generating…' : 'Generate'}
          </button>
        </div>
      </div>
    </div>
  );
}

const inputClass =
  'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500';

function FieldGroup({
  label,
  required,
  children,
}: {
  label: string;
  required?: boolean;
  children: React.ReactNode;
}) {
  return (
    <div>
      <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
        {label}
        {required && <span className="text-error ml-0.5">*</span>}
      </label>
      {children}
    </div>
  );
}
