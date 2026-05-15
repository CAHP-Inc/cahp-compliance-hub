import { useState, useRef, DragEvent } from 'react';
import { uploadDocument } from '../lib/sharepoint';
import { Icon } from './ui/Icon';

/**
 * The 8 property-linked SharePoint Document Libraries. Each has a PropertyLookupId metadata column.
 * Picked from PropertyDocumentsTab — same list, single source of truth.
 */
export const PROPERTY_LINKED_LIBRARIES = [
  'AMI Certification Letters',
  'DOR Correspondence',
  'DOR Submittal Packages',
  'Land Use Restriction Agreements',
  'Operating Agreements',
  'Org Charts',
  'Property Deeds',
  'Supporting Documentation',
] as const;

export type PropertyLinkedLibrary = (typeof PROPERTY_LINKED_LIBRARIES)[number];

/**
 * The dedicated CAHP Entity Documents library. Every file here is a CAHP entity
 * document — the library IS the filter, no OwnerLookupId tagging required.
 *
 * Used by the CAHP Entity page and surfaced as "CAHP Entity Reference Documents"
 * at the top of every property's Documents tab.
 */
export const CAHP_ENTITY_LIBRARY = 'CAHP Entity Documents';

export type UploadScope =
  | { type: 'property'; propertyId: string; propertyTitle?: string }
  | { type: 'owner'; ownerId: string; ownerTitle?: string };

/**
 * Suggested library per property document context — used to pre-select the dropdown
 * when the modal is opened from a workflow with a hint (e.g., the Log Letter modal
 * always wants "DOR Correspondence").
 */
export interface UploadDocumentModalProps {
  /** Modern API — pass a scope. Property or owner. */
  scope?: UploadScope;
  /** Legacy API — kept for backward compat with existing call sites */
  propertyId?: string;
  propertyTitle?: string;
  /** Default library to pre-select. If omitted, "Supporting Documentation" is used. */
  defaultLibrary?: PropertyLinkedLibrary;
  /** Called after a successful upload with the result */
  onSuccess: (result: { filename: string; webUrl: string; library: string }) => void;
  onClose: () => void;
}

export function UploadDocumentModal({
  scope,
  propertyId: legacyPropertyId,
  propertyTitle: legacyPropertyTitle,
  defaultLibrary,
  onSuccess,
  onClose,
}: UploadDocumentModalProps) {
  // Normalize: if legacy props were used, build a property scope
  const effectiveScope: UploadScope = scope ?? {
    type: 'property',
    propertyId: legacyPropertyId!,
    propertyTitle: legacyPropertyTitle,
  };

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [library, setLibrary] = useState<PropertyLinkedLibrary>(
    defaultLibrary ?? 'Supporting Documentation'
  );
  const [dragOver, setDragOver] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);

  const handleFileSelect = (selectedFile: File) => {
    setFile(selectedFile);
    setError(null);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragOver(false);
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile) handleFileSelect(droppedFile);
  };

  const handleUpload = async () => {
    if (!file) {
      setError('Pick a file first.');
      return;
    }
    setUploading(true);
    setError(null);
    setProgress(0);

    try {
      const metadata =
        effectiveScope.type === 'property'
          ? { PropertyLookupId: effectiveScope.propertyId }
          : { OwnerLookupId: effectiveScope.ownerId };

      const result = await uploadDocument({
        libraryName: library,
        filename: file.name,
        file,
        metadata,
        onProgress: setProgress,
      });

      onSuccess({
        filename: result.filename,
        webUrl: result.webUrl,
        library,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setUploading(false);
    }
  };

  const scopeLabel =
    effectiveScope.type === 'property'
      ? effectiveScope.propertyTitle
      : effectiveScope.ownerTitle;
  const metadataLabel = effectiveScope.type === 'property' ? 'PropertyID' : 'OwnerID';

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-md w-full p-5">
        <h3 className="text-lg font-bold text-teal-700 mb-1">Upload Document</h3>
        <p className="text-sm text-gray-600 mb-4">
          {scopeLabel ? (
            <>
              Uploading to <strong>{scopeLabel}</strong>. File goes into the SharePoint library you pick;
              {' '}{metadataLabel} metadata is set automatically so it surfaces in the right places.
            </>
          ) : (
            <>File goes into the SharePoint library you pick; {metadataLabel} is set automatically.</>
          )}
        </p>

        {/* Drop zone */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
          onClick={() => !uploading && fileInputRef.current?.click()}
          className={`border-2 border-dashed rounded-lg p-6 text-center cursor-pointer transition-colors ${
            dragOver
              ? 'border-teal-500 bg-teal-50'
              : file
                ? 'border-success bg-green-50'
                : 'border-gray-300 hover:border-teal-400 hover:bg-gray-50'
          } ${uploading ? 'opacity-60 cursor-not-allowed' : ''}`}
        >
          <input
            ref={fileInputRef}
            type="file"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) handleFileSelect(f);
            }}
            disabled={uploading}
          />
          {file ? (
            <div className="text-sm">
              <Icon name="check" size={24} className="text-success mx-auto mb-2" />
              <div className="font-medium text-gray-900 break-all">{file.name}</div>
              <div className="text-xs text-gray-500 mt-1 font-mono-data">
                {formatFileSize(file.size)}
                {file.size > 4 * 1024 * 1024 && (
                  <span className="ml-2 text-teal-700">· chunked upload</span>
                )}
              </div>
              {!uploading && (
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setFile(null);
                    setError(null);
                  }}
                  className="text-xs text-teal-700 hover:text-teal-900 mt-2 underline"
                >
                  Choose a different file
                </button>
              )}
            </div>
          ) : (
            <div className="text-sm text-gray-500">
              <Icon name="folder" size={24} className="text-gray-400 mx-auto mb-2" />
              <p className="font-medium text-gray-700">Drop a file here, or click to browse</p>
              <p className="text-xs text-gray-400 mt-1">Any file type · large files upload in chunks with progress</p>
            </div>
          )}
        </div>

        {/* Library picker */}
        <div className="mt-4">
          <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wider mb-1">
            SharePoint Library
          </label>
          <select
            value={library}
            onChange={(e) => setLibrary(e.target.value as PropertyLinkedLibrary)}
            disabled={uploading}
            className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 bg-white disabled:opacity-50"
          >
            {PROPERTY_LINKED_LIBRARIES.map((lib) => (
              <option key={lib} value={lib}>{lib}</option>
            ))}
          </select>
          <p className="text-[11px] text-gray-400 mt-0.5">
            Documents are organized by category, not by property. PropertyID metadata associates each file with this property.
          </p>
        </div>

        {/* Progress */}
        {uploading && (
          <div className="mt-4">
            <div className="flex items-center justify-between text-xs text-gray-600 mb-1">
              <span>Uploading…</span>
              <span className="font-mono-data">{progress}%</span>
            </div>
            <div className="h-1.5 bg-gray-100 rounded-full overflow-hidden">
              <div
                className="h-full bg-teal-500 transition-all"
                style={{ width: `${progress}%` }}
              />
            </div>
          </div>
        )}

        {/* Error */}
        {error && (
          <div className="mt-3 bg-red-50 border border-red-200 rounded-md p-2 text-xs text-red-800">
            {error}
          </div>
        )}

        {/* Actions */}
        <div className="flex items-center justify-end gap-2 mt-5">
          <button
            onClick={onClose}
            disabled={uploading}
            className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
          >
            Cancel
          </button>
          <button
            onClick={handleUpload}
            disabled={uploading || !file}
            className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {uploading && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {uploading ? 'Uploading…' : 'Upload'}
          </button>
        </div>
      </div>
    </div>
  );
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`;
}
