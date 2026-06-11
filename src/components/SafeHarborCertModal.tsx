import { useMemo, useState, useRef, type DragEvent } from 'react';
import {
  useSharePointList,
  uploadDocument,
  getUpstreamOwnerIds,
  LIST_NAMES,
  type Owner,
  type Ownership,
  type Property,
} from '../lib/sharepoint';
import {
  analyze,
  buildExhibitBlob,
  buildLetterDocx,
  buildLetterPdf,
  defaultCertConfig,
  deriveCertConfig,
  parseRentRoll,
  slugify,
  type CertConfig,
  type ParsedRoll,
  type Unit,
} from '../lib/safe-harbor';
import { Icon } from './ui/Icon';

const CERT_LIBRARY = 'AMI Certification Letters';

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export interface SafeHarborCertModalProps {
  /** Optional pre-selected property (e.g. launched from a property page). */
  initialPropertyId?: string;
  onClose: () => void;
}

export function SafeHarborCertModal({ initialPropertyId, onClose }: SafeHarborCertModalProps) {
  const properties = useSharePointList<Property>(LIST_NAMES.Properties, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const loading = properties.loading || owners.loading || ownership.loading;

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [rolls, setRolls] = useState<ParsedRoll[]>([]);
  const [parseError, setParseError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);

  // Selection value is "prop:<id>" (single LLC) or "owner:<id>" (portfolio parent).
  const [entitySel, setEntitySel] = useState(initialPropertyId ? `prop:${initialPropertyId}` : '');
  const [taxYear, setTaxYear] = useState(new Date().getFullYear());
  const [utilityAllowance, setUtilityAllowance] = useState(0);
  const [relationship, setRelationship] = useState('property manager and authorized agent');
  const [description, setDescription] = useState('scattered-site residential rental units');
  const [parcels, setParcels] = useState('');
  const [groupOverride, setGroupOverride] = useState(false);
  const [groupName, setGroupName] = useState('');

  const [busy, setBusy] = useState<string | null>(null);
  const [saveMsg, setSaveMsg] = useState<string | null>(null);
  const [saveErr, setSaveErr] = useState<string | null>(null);

  const handleFiles = async (fileList: FileList | File[]) => {
    setParseError(null);
    const parsed: ParsedRoll[] = [];
    for (const f of Array.from(fileList)) {
      try {
        parsed.push(await parseRentRoll(f));
      } catch (e) {
        setParseError(e instanceof Error ? e.message : String(e));
      }
    }
    if (parsed.length) {
      setRolls((prev) => [...prev, ...parsed]);
      if (!groupName && parsed[0]) setGroupName('');
    }
  };
  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files?.length) void handleFiles(e.dataTransfer.files);
  };
  const removeRoll = (i: number) => setRolls((prev) => prev.filter((_, idx) => idx !== i));

  const [selType, selId] = useMemo(() => {
    const i = entitySel.indexOf(':');
    return i === -1 ? ['', ''] : [entitySel.slice(0, i), entitySel.slice(i + 1)];
  }, [entitySel]);

  const ownerSelected = useMemo(
    () => (selType === 'owner' ? owners.data?.find((o) => String(o.id) === selId) : undefined),
    [selType, selId, owners.data],
  );

  // Child properties of a selected owner (one row per LLC the owner holds).
  const childProperties = useMemo(() => {
    if (selType !== 'owner' || !ownership.data || !properties.data) return [];
    const childIds = new Set(
      ownership.data
        .filter((r) => String(r.fields.OwnerLookupId) === selId)
        .map((r) => String(r.fields.LinkedPropertyLookupId)),
    );
    return properties.data.filter((p) => childIds.has(String(p.id)));
  }, [selType, selId, ownership.data, properties.data]);

  // The property we derive CAHP/EIN boilerplate from: the chosen LLC, or a
  // representative child of the chosen owner (prefer one with a CAHP chain).
  const representativeProperty = useMemo(() => {
    if (selType === 'prop') return properties.data?.find((p) => String(p.id) === selId);
    if (selType === 'owner' && childProperties.length && ownership.data) {
      const withCahp = childProperties.find((p) => {
        const up = getUpstreamOwnerIds(String(p.id), ownership.data!);
        return [...up].some((id) => owners.data?.find((o) => String(o.id) === id)?.fields.IsCAHPEntity);
      });
      return withCahp ?? childProperties[0];
    }
    return undefined;
  }, [selType, selId, properties.data, childProperties, ownership.data, owners.data]);

  // Derive boilerplate (falls back to defaults for an owner with no linked LLCs).
  const derived = useMemo(() => {
    if (representativeProperty && owners.data && ownership.data) {
      const upstream = getUpstreamOwnerIds(String(representativeProperty.id), ownership.data);
      return deriveCertConfig({
        property: representativeProperty as unknown as Parameters<typeof deriveCertConfig>[0]['property'],
        owners: owners.data as unknown as Parameters<typeof deriveCertConfig>[0]['owners'],
        ownership: ownership.data as unknown as Parameters<typeof deriveCertConfig>[0]['ownership'],
        upstreamOwnerIds: upstream,
        taxYear,
      });
    }
    if (ownerSelected) {
      return { config: defaultCertConfig(ownerSelected.fields.Title ?? '', taxYear), exemptionChainOk: true, warnings: [] };
    }
    return null;
  }, [representativeProperty, ownerSelected, owners.data, ownership.data, taxYear]);

  const units: Unit[] = useMemo(() => rolls.flatMap((r) => r.units), [rolls]);
  const distinctSources = useMemo(
    () => [...new Set(units.map((u) => u.source).filter(Boolean))],
    [units],
  );
  const isGroup = Boolean(ownerSelected) || groupOverride || distinctSources.length > 1;

  const analysis = useMemo(
    () => (units.length ? analyze(units, { taxYear, utilityAllowance, forceGroup: Boolean(ownerSelected) || groupOverride }) : null),
    [units, taxYear, utilityAllowance, ownerSelected, groupOverride],
  );

  // Final config = derived hub facts + UI overrides.
  const config: CertConfig | null = useMemo(() => {
    if (!derived) return null;
    const c: CertConfig = JSON.parse(JSON.stringify(derived.config));
    c.certification.relationshipToOwner = relationship;
    c.property.description = description;
    c.property.taxMapParcels = parcels.split(',').map((s) => s.trim()).filter(Boolean);
    c.filing.taxYear = taxYear;
    // For a group (or when the hub had no county), take counties from the rent rolls.
    const rollCounties = [...new Set(units.map((u) => u.county).filter((x): x is string => Boolean(x)))];
    if ((isGroup || c.property.counties.length === 0) && rollCounties.length) {
      c.property.counties = rollCounties;
      if (!c.property.addressLine) {
        c.property.addressLine = `Scattered sites located in ${rollCounties.join(' and ')} ${rollCounties.length > 1 ? 'Counties' : 'County'}, South Carolina`;
      }
    }
    if (isGroup) {
      const gName = groupName || ownerSelected?.fields.Title || derived.config.company.legalName;
      if (ownerSelected?.fields.Title) c.company.legalName = ownerSelected.fields.Title;
      c.portfolio = {
        isGroupFiling: true,
        groupName: gName,
        groupStateType: 'South Carolina limited liability company',
        subsidiaryDescription: 'wholly-owned single-purpose subsidiary LLCs',
      };
    }
    return c;
  }, [derived, relationship, description, parcels, taxYear, isGroup, groupName, ownerSelected, units]);

  const ready = Boolean(analysis && config);
  const baseName = useMemo(() => {
    if (!config) return 'Safe_Harbor';
    return isGroup
      ? `${slugify(config.portfolio!.groupName)}_GROUP`
      : slugify(config.company.legalName);
  }, [config, isGroup]);

  const doDownload = async (kind: 'docx' | 'pdf' | 'xlsx') => {
    if (!analysis || !config) return;
    setBusy(kind);
    try {
      if (kind === 'xlsx') {
        downloadBlob(buildExhibitBlob(analysis, config), `${baseName}_Exhibit_A_Unit_AMI_Analysis.xlsx`);
      } else if (kind === 'pdf') {
        downloadBlob(buildLetterPdf(analysis, config), `${baseName}_Safe_Harbor_Certification_TY${taxYear}.pdf`);
      } else {
        downloadBlob(await buildLetterDocx(analysis, config), `${baseName}_Safe_Harbor_Certification_TY${taxYear}.docx`);
      }
    } finally {
      setBusy(null);
    }
  };

  const saveToEntity = async () => {
    if (!analysis || !config || !representativeProperty) return;
    setBusy('save');
    setSaveErr(null);
    setSaveMsg(null);
    try {
      const docxBlob = await buildLetterDocx(analysis, config);
      const xlsxBlob = buildExhibitBlob(analysis, config);
      const meta = { PropertyLookupId: String(representativeProperty.id) };
      await uploadDocument({
        libraryName: CERT_LIBRARY,
        filename: `${baseName}_Safe_Harbor_Certification_TY${taxYear}.docx`,
        file: docxBlob,
        metadata: meta,
      });
      await uploadDocument({
        libraryName: CERT_LIBRARY,
        filename: `${baseName}_Exhibit_A_Unit_AMI_Analysis.xlsx`,
        file: xlsxBlob,
        metadata: meta,
      });
      setSaveMsg(
        `Saved letter + Exhibit A to "${CERT_LIBRARY}", tagged to ${config.company.legalName}. ` +
          `They now appear on that property's Documents tab.`,
      );
    } catch (e) {
      setSaveErr(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(null);
    }
  };

  const sc = analysis?.scopes;
  const p = analysis?.roll.pct;
  const cnt = analysis?.roll.counts;

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full p-5 my-8">
        <div className="flex items-start justify-between mb-1">
          <h3 className="text-lg font-bold text-teal-700">Generate Safe Harbor Certification</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600"><Icon name="x" size={18} /></button>
        </div>
        <p className="text-sm text-gray-600 mb-4">
          Upload AppFolio rent roll(s), pick the filing entity, and generate the §12-37-220(B)(11)(e) /
          Rev. Proc. 96-32 certification letter (.docx + PDF) and Exhibit A. Qualification is rent-based;
          drop more than one rent roll to file a portfolio (group) of LLCs together.
        </p>

        {loading ? (
          <div className="py-8 text-center text-sm text-gray-500">Loading entities…</div>
        ) : (
          <div className="space-y-4">
            {/* Rent roll dropzone */}
            <div
              onDrop={onDrop}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onClick={() => fileInputRef.current?.click()}
              className={`border-2 border-dashed rounded-md p-4 text-center cursor-pointer text-sm ${
                dragOver ? 'border-teal-500 bg-teal-50' : 'border-gray-300 hover:bg-gray-50'
              }`}
            >
              <input
                ref={fileInputRef} type="file" accept=".xlsx" multiple className="hidden"
                onChange={(e) => e.target.files && handleFiles(e.target.files)}
              />
              <Icon name="file" size={18} className="mx-auto mb-1 text-gray-400" />
              Drop AppFolio rent roll .xlsx file(s) here, or click to browse
            </div>
            {parseError && <div className="bg-red-50 border border-red-200 rounded p-2 text-xs text-red-800">{parseError}</div>}
            {rolls.length > 0 && (
              <div className="text-xs space-y-1">
                {rolls.map((r, i) => (
                  <div key={i} className="flex items-center justify-between bg-gray-50 border border-gray-200 rounded px-2 py-1">
                    <span><strong>{r.source}</strong> — {r.units.length} units <span className="text-gray-400">({r.filename})</span></span>
                    <button onClick={() => removeRoll(i)} className="text-gray-400 hover:text-red-600"><Icon name="trash" size={12} /></button>
                  </div>
                ))}
              </div>
            )}

            {/* Entity + options */}
            <div className="grid grid-cols-2 gap-3 text-sm">
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-gray-600">Filing entity</span>
                <select value={entitySel} onChange={(e) => setEntitySel(e.target.value)} className="border border-gray-300 rounded px-2 py-1">
                  <option value="">— select entity —</option>
                  <optgroup label="Portfolio / parent owner (group filing)">
                    {(owners.data ?? []).slice().sort((a, b) => (a.fields.Title || '').localeCompare(b.fields.Title || '')).map((o) => (
                      <option key={`o${o.id}`} value={`owner:${o.id}`}>{o.fields.Title}</option>
                    ))}
                  </optgroup>
                  <optgroup label="Single LLC (property)">
                    {(properties.data ?? []).slice().sort((a, b) => (a.fields.Title || '').localeCompare(b.fields.Title || '')).map((pr) => (
                      <option key={`p${pr.id}`} value={`prop:${pr.id}`}>{pr.fields.Title}</option>
                    ))}
                  </optgroup>
                </select>
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-gray-600">Tax year</span>
                <input type="number" value={taxYear} onChange={(e) => setTaxYear(Number(e.target.value))} className="border border-gray-300 rounded px-2 py-1" />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-gray-600">Utility allowance ($, added to contract rent)</span>
                <input type="number" value={utilityAllowance} onChange={(e) => setUtilityAllowance(Number(e.target.value))} className="border border-gray-300 rounded px-2 py-1" />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-gray-600">Certifier relationship to owner</span>
                <input value={relationship} onChange={(e) => setRelationship(e.target.value)} className="border border-gray-300 rounded px-2 py-1" />
              </label>
              <label className="flex flex-col gap-1 col-span-2">
                <span className="text-xs font-semibold text-gray-600">Property description</span>
                <input value={description} onChange={(e) => setDescription(e.target.value)} className="border border-gray-300 rounded px-2 py-1" />
              </label>
              <label className="flex flex-col gap-1 col-span-2">
                <span className="text-xs font-semibold text-gray-600">Tax Map / Parcel No(s) — comma separated</span>
                <input value={parcels} onChange={(e) => setParcels(e.target.value)} placeholder="optional" className="border border-gray-300 rounded px-2 py-1" />
              </label>
              <label className="flex items-center gap-2 col-span-2 text-xs text-gray-700">
                <input type="checkbox" checked={isGroup} disabled={distinctSources.length > 1 || Boolean(ownerSelected)} onChange={(e) => setGroupOverride(e.target.checked)} />
                Group / portfolio filing
                {ownerSelected
                  ? <span className="text-gray-400">(auto: parent owner selected)</span>
                  : distinctSources.length > 1 && <span className="text-gray-400">(auto: {distinctSources.length} LLCs detected)</span>}
              </label>
              {isGroup && (
                <label className="flex flex-col gap-1 col-span-2">
                  <span className="text-xs font-semibold text-gray-600">Portfolio (parent) name</span>
                  <input value={groupName || ownerSelected?.fields.Title || ''} onChange={(e) => setGroupName(e.target.value)} placeholder="e.g. IV Fund Global, LLC" className="border border-gray-300 rounded px-2 py-1" />
                </label>
              )}
            </div>

            {/* Live determination preview */}
            {analysis && p && cnt && sc && (
              <div className="border border-gray-200 rounded-md p-3">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-sm font-bold text-gray-800">Determination</div>
                  <div className={`text-sm font-bold ${sc.qualifies.length ? 'text-success' : 'text-warning'}`}>{sc.headline}</div>
                </div>
                <div className="text-xs text-gray-600 mb-2">
                  {analysis.roll.denom} residential units{analysis.roll.nNonRes ? ` · ${analysis.roll.nNonRes} non-residential excluded` : ''}
                  {analysis.roll.nReview ? ` · ${analysis.roll.nReview} need review` : ''}
                  {isGroup ? ` · ${distinctSources.length} LLCs` : ''}
                </div>
                <table className="w-full text-xs">
                  <thead className="text-[10px] text-gray-500 uppercase">
                    <tr><th className="text-left">AMI Tier</th><th>Units</th><th>%</th><th>Required</th><th>Result</th></tr>
                  </thead>
                  <tbody>
                    {[
                      ['Low-Income (≤80%)', cnt.le80, p.le80, '≥75%', p.le80 >= 75],
                      ['≤60% (40/60 scope)', cnt.le60, p.le60, '≥40%', p.le60 >= 40],
                      ['Very Low (≤50%) (20/50 scope)', cnt.le50, p.le50, '≥20%', p.le50 >= 20],
                      ['Market (>80%)', cnt.market, p.market, '≤25%', p.market <= 25],
                    ].map(([label, u, pc, req, ok], i) => (
                      <tr key={i} className="border-t border-gray-100">
                        <td className="py-0.5">{label as string}</td>
                        <td className="text-center">{u as number}</td>
                        <td className="text-center">{pc as number}%</td>
                        <td className="text-center">{req as string}</td>
                        <td className={`text-center font-semibold ${ok ? 'text-success' : 'text-warning'}`}>{ok ? 'PASS' : 'FAIL'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {analysis.roll.nReview > 0 && (
                  <div className="mt-2 text-[11px] text-amber-700">
                    {analysis.roll.nReview} unit(s) need review (missing bedroom/rent/county) — see Exhibit A. Percentages treat them conservatively.
                  </div>
                )}
              </div>
            )}

            {saveMsg && <div className="bg-green-50 border border-green-200 rounded p-2 text-xs text-green-800">{saveMsg}</div>}
            {saveErr && <div className="bg-red-50 border border-red-200 rounded p-2 text-xs text-red-800">{saveErr}</div>}
          </div>
        )}

        <div className="flex items-center justify-between gap-2 mt-5">
          <button onClick={onClose} className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm">Close</button>
          <div className="flex items-center gap-2">
            <button disabled={!ready || busy !== null} onClick={() => doDownload('docx')} className="px-3 py-1.5 border border-teal-300 text-teal-700 hover:bg-teal-50 rounded-md text-sm font-medium disabled:opacity-50">{busy === 'docx' ? '…' : 'Letter .docx'}</button>
            <button disabled={!ready || busy !== null} onClick={() => doDownload('pdf')} className="px-3 py-1.5 border border-teal-300 text-teal-700 hover:bg-teal-50 rounded-md text-sm font-medium disabled:opacity-50">{busy === 'pdf' ? '…' : 'Letter PDF'}</button>
            <button disabled={!ready || busy !== null} onClick={() => doDownload('xlsx')} className="px-3 py-1.5 border border-teal-300 text-teal-700 hover:bg-teal-50 rounded-md text-sm font-medium disabled:opacity-50">{busy === 'xlsx' ? '…' : 'Exhibit .xlsx'}</button>
            <button disabled={!ready || !representativeProperty || busy !== null} title={!representativeProperty ? 'Pick a property (or an owner with at least one linked LLC) to save into the hub' : ''} onClick={saveToEntity} className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium disabled:opacity-50">{busy === 'save' ? 'Saving…' : 'Save to entity'}</button>
          </div>
        </div>
      </div>
    </div>
  );
}
