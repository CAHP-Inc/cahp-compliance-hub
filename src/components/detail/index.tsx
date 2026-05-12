/**
 * Reusable building blocks for module detail pages with inline editing.
 *
 * Used by: PropertyDetail, ComplianceDeadlineDetail, and future detail pages
 * (Submittals, Outstanding Items, etc.).
 *
 * Architecture: each detail page tracks its own draft state and save logic;
 * these components provide the visual primitives (sections, breadcrumb,
 * editable form field that swaps between read/edit modes).
 */
import { Link } from 'react-router-dom';

// =============================================================================
// BreadcrumbBar
// =============================================================================

export function BreadcrumbBar({
  parentLabel,
  parentTo,
  currentLabel,
}: {
  parentLabel: string;
  parentTo: string;
  currentLabel?: string;
}) {
  return (
    <nav className="mb-4 text-sm">
      <Link to={parentTo} className="text-teal-700 hover:text-teal-900 font-medium">
        ← {parentLabel}
      </Link>
      {currentLabel && (
        <>
          <span className="mx-2 text-gray-400">/</span>
          <span className="text-gray-700">{currentLabel}</span>
        </>
      )}
    </nav>
  );
}

// =============================================================================
// Section card wrapper
// =============================================================================

export function Section({
  title,
  children,
  fullWidth,
}: {
  title: string;
  children: React.ReactNode;
  fullWidth?: boolean;
}) {
  return (
    <div
      className={`bg-white border border-gray-200 rounded-lg p-5 shadow-card ${
        fullWidth ? 'lg:col-span-2' : ''
      }`}
    >
      <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">
        {title}
      </h3>
      <div className="space-y-2.5">{children}</div>
    </div>
  );
}

// =============================================================================
// EditableField — swaps between read and edit mode
// =============================================================================

export type FieldValue = string | number | null | undefined;
export type FieldType = 'text' | 'textarea' | 'number' | 'date' | 'choice';

export function EditableField({
  label,
  value,
  editing,
  onChange,
  type = 'text',
  choices,
  mono,
  required,
  rows = 3,
  hideLabel,
}: {
  label: string;
  value: FieldValue;
  editing: boolean;
  onChange: (v: FieldValue) => void;
  type?: FieldType;
  choices?: readonly string[];
  mono?: boolean;
  required?: boolean;
  rows?: number;
  hideLabel?: boolean;
}) {
  if (!editing) {
    let displayValue: React.ReactNode = value;
    if (type === 'date' && value) displayValue = formatDate(value as string);
    if (type === 'textarea' && value) {
      displayValue = (
        <span className="whitespace-pre-wrap leading-relaxed block">{String(value)}</span>
      );
    }
    return (
      <div className="flex items-start gap-3">
        {!hideLabel && <dt className="text-sm text-gray-500 w-44 flex-shrink-0">{label}</dt>}
        <dd className={`text-sm text-gray-900 flex-1 ${mono ? 'font-mono-data' : ''}`}>
          {value === null || value === undefined || value === '' ? (
            <span className="text-gray-300">—</span>
          ) : (
            displayValue
          )}
        </dd>
      </div>
    );
  }

  const inputClass =
    'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  let input: React.ReactNode;
  if (type === 'textarea') {
    input = (
      <textarea
        value={String(value ?? '')}
        onChange={(e) => onChange(e.target.value)}
        rows={rows}
        className={`${inputClass} resize-y`}
      />
    );
  } else if (type === 'number') {
    input = (
      <input
        type="number"
        value={value == null || value === '' ? '' : value}
        onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}
        className={`${inputClass} ${mono ? 'font-mono-data' : ''}`}
      />
    );
  } else if (type === 'date') {
    const dateValue = value ? new Date(value as string).toISOString().slice(0, 10) : '';
    input = (
      <input
        type="date"
        value={dateValue}
        onChange={(e) =>
          onChange(e.target.value ? new Date(e.target.value).toISOString() : null)
        }
        className={`${inputClass} font-mono-data`}
      />
    );
  } else if (type === 'choice') {
    input = (
      <select
        value={String(value ?? '')}
        onChange={(e) => onChange(e.target.value || null)}
        className={`${inputClass} bg-white`}
      >
        <option value="">— none —</option>
        {choices?.map((c) => (
          <option key={c} value={c}>
            {c}
          </option>
        ))}
      </select>
    );
  } else {
    input = (
      <input
        type="text"
        value={String(value ?? '')}
        onChange={(e) => onChange(e.target.value)}
        className={`${inputClass} ${mono ? 'font-mono-data' : ''}`}
        required={required}
      />
    );
  }

  return (
    <div className="flex items-start gap-3">
      {!hideLabel && (
        <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
          {label}
          {required && <span className="text-error ml-0.5">*</span>}
        </label>
      )}
      <div className="flex-1">{input}</div>
    </div>
  );
}

// =============================================================================
// SaveErrorBanner — consistent error display for failed saves
// =============================================================================

import { Icon } from '../ui/Icon';

export function SaveErrorBanner({ error }: { error: string | null }) {
  if (!error) return null;
  return (
    <div className="mb-4 bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
      <Icon name="alert" size={16} className="text-error flex-shrink-0 mt-0.5" />
      <div className="flex-1 text-sm">
        <span className="font-semibold text-error">Save failed.</span>{' '}
        <span className="text-red-700 font-mono-data text-xs">{error}</span>
      </div>
    </div>
  );
}

// =============================================================================
// EditingActionButtons — Save/Cancel pair shown when in edit mode
// =============================================================================

export function EditingActionButtons({
  saving,
  onSave,
  onCancel,
}: {
  saving: boolean;
  onSave: () => void;
  onCancel: () => void;
}) {
  return (
    <>
      <button
        onClick={onCancel}
        disabled={saving}
        className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
      >
        Cancel
      </button>
      <button
        onClick={onSave}
        disabled={saving}
        className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
      >
        {saving && (
          <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
        )}
        {saving ? 'Saving…' : 'Save Changes'}
      </button>
    </>
  );
}

// =============================================================================
// formatDate utility — exported so detail pages can format dates consistently
// =============================================================================

export function formatDate(iso: string | undefined | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
