import { useCallback, useMemo, useState, useRef, type DragEvent } from 'react';
import {
  useSharePointList,
  uploadDocument,
  getUpstreamOwnerIds,
  isCahpEntity,
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
  countiesForState,
  defaultCertConfig,
  deriveCertConfig,
  jurisdictionDefaults,
  extractPdfText,
  parseRentRoll,
  parseTextToUnits,
  slugify,
  type CertConfig,
  type ParsedRoll,
  type Unit,
} from '../lib/safe-harbor';
import { Icon } from './ui/Icon';

const CERT_LIBRARY = 'AMI Certification Letters';

// Beneficial % held by a CAHP entity, walking UP the ownership chain and
// multiplying percentages. Unlike the org-chart engine (which records only
// terminal Individuals/Nonprofits), this STOPS at the first CAHP-flagged entity
// on each path — so a CAHP LLC instrumentality (e.g. "CAHP SC, LLC") counts even
// when nothing terminal sits above it. Returns null if no CAHP entity is found.
function cahpBeneficialPercent(
  subjectType: 'property' | 'owner',
  subjectId: string,
  owners: Owner[],
  ownership: Ownership[],
): number | null {
  const ownerById = new Map(owners.map((o) => [String(o.id), o]));
  let total = 0;
  let found = false;
  const walk = (type: 'property' | 'owner', id: string, cumPct: number, visited: Set<string>) => {
    const rows =
      type === 'property'
        ? ownership.filter((r) => String(r.fields.LinkedPropertyLookupId) === id)
        : ownership.filter((r) => String(r.fields.ParentOwnerLookupId) === id);
    for (const r of rows) {
      const ownerId = r.fields.OwnerLookupId ? String(r.fields.OwnerLookupId) : '';
      if (!ownerId || visited.has(ownerId)) continue;
      const owner = ownerById.get(ownerId);
      const eff = (cumPct * (r.fields.OwnershipPercent ?? 0)) / 100;
      if (owner && isCahpEntity(owner)) {
        total += eff;
        found = true; // CAHP is the endpoint we care about — don't recurse past it
      } else {
        walk('owner', ownerId, eff, new Set(visited).add(ownerId));
      }
    }
  };
  walk(subjectType, subjectId, 100, new Set());
  return found ? Math.round(total * 10000) / 10000 : null;
}

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

  // Non-AppFolio input (paste text or a PDF rent roll from an outside PM).
  const [inputMode, setInputMode] = useState<'appfolio' | 'paste'>('appfolio');
  const [pasteText, setPasteText] = useState('');
  const [pasteSource, setPasteSource] = useState('');
  const [pasteCounty, setPasteCounty] = useState('');
  const [pasteBedrooms, setPasteBedrooms] = useState('');
  const [pdfBusy, setPdfBusy] = useState(false);
  // State for prospect / paste filings (hub properties carry their own cahpState).
  const [manualState, setManualState] = useState<'SC' | 'NC'>('SC');

  // Selection value is "prop:<id>" (single LLC) or "owner:<id>" (portfolio parent).
  const [entitySel, setEntitySel] = useState(initialPropertyId ? `prop:${initialPropertyId}` : '');
  const [taxYear, setTaxYear] = useState(new Date().getFullYear());
  const [utilityAllowance, setUtilityAllowance] = useState(0);
  const [relationship, setRelationship] = useState('property manager and authorized agent');
  const [citationOverride, setCitationOverride] = useState('');
  const [recipientOverride, setRecipientOverride] = useState('');
  const [description, setDescription] = useState('scattered-site residential rental units');
  const [groupOverride, setGroupOverride] = useState(false);
  const [groupName, setGroupName] = useState('');
  // Prospect mode: entity isn't in the hub yet — qualification only.
  const [prospectMode, setProspectMode] = useState(false);
  const [prospectName, setProspectName] = useState('');

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

  const handlePdfFile = async (file: File) => {
    setParseError(null);
    setPdfBusy(true);
    try {
      const text = await extractPdfText(file);
      setPasteText((prev) => (prev ? prev + '\n' : '') + text);
      if (!pasteSource) setPasteSource(file.name.replace(/\.pdf$/i, ''));
    } catch (e) {
      setParseError(
        'Could not read that PDF automatically (' + (e instanceof Error ? e.message : String(e)) +
          '). Open it, copy the text, and paste it into the box instead.',
      );
    } finally {
      setPdfBusy(false);
    }
  };

  const handleParseText = () => {
    setParseError(null);
    const source = pasteSource.trim() || 'Pasted rent roll';
    const parsed = parseTextToUnits(pasteText, {
      source,
      defaultCounty: pasteCounty || null,
      defaultBedrooms: pasteBedrooms === '' ? null : Number(pasteBedrooms),
    });
    if (!parsed.length) {
      setParseError('No units found. Each unit needs a line like "Unit 1 … $995" — check the text or enter rows by editing the source.');
      return;
    }
    setRolls((prev) => [...prev, { units: parsed, exported: '', source, filename: 'pasted/PDF' }]);
    setPasteText('');
  };

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

  // Filing state: a hub property carries its own cahpState; prospect/paste use the manual selector.
  const filingState = ((representativeProperty?.fields.cahpState as string | undefined) || manualState || 'SC').toUpperCase();

  // Derive boilerplate (falls back to defaults for an owner with no linked LLCs).
  const derived = useMemo(() => {
    if (prospectMode) {
      return { config: defaultCertConfig(prospectName || 'Prospective Entity', taxYear, filingState), exemptionChainOk: true, warnings: [] };
    }
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
      return { config: defaultCertConfig(ownerSelected.fields.Title ?? '', taxYear, filingState), exemptionChainOk: true, warnings: [] };
    }
    return null;
  }, [prospectMode, prospectName, representativeProperty, ownerSelected, owners.data, ownership.data, taxYear, filingState]);

  // CAHP's beneficial ownership % of the SELECTED PARENT fund (computed once).
  const parentCahpPct = useMemo(() => {
    if (!ownerSelected || !owners.data || !ownership.data) return null;
    return cahpBeneficialPercent('owner', String(ownerSelected.id), owners.data, ownership.data);
  }, [ownerSelected, owners.data, ownership.data]);

  // CAHP nonprofit's beneficial ownership %/class for a property (compounded up the
  // chain; for a parent-fund filing it flows from the parent × its stake, default
  // 100%). Member class comes from a direct CAHP member row when present.
  const cahpInterestFor = useCallback(
    (propertyId: string): { ownershipPercent: number | null; memberClass: string } => {
      if (!owners.data || !ownership.data) return { ownershipPercent: null, memberClass: '' };
      const round = (n: number) => Math.round(n * 10000) / 10000;
      const ownerById = new Map(owners.data.map((o) => [String(o.id), o]));
      const directRow = ownership.data.find(
        (r) =>
          String(r.fields.LinkedPropertyLookupId) === propertyId &&
          ownerById.get(String(r.fields.OwnerLookupId))?.fields.IsCAHPEntity,
      );
      const memberClass = directRow?.fields.MemberClass || '';
      if (ownerSelected && parentCahpPct != null) {
        const parentOwnPct =
          ownership.data.find(
            (r) =>
              String(r.fields.LinkedPropertyLookupId) === propertyId &&
              String(r.fields.OwnerLookupId) === String(ownerSelected.id),
          )?.fields.OwnershipPercent ?? 100;
        return { ownershipPercent: round((parentCahpPct * parentOwnPct) / 100), memberClass };
      }
      const pct = cahpBeneficialPercent('property', propertyId, owners.data, ownership.data);
      return { ownershipPercent: pct, memberClass };
    },
    [owners.data, ownership.data, ownerSelected, parentCahpPct],
  );

  // Units = all rent-roll units, aggregated (no per-unit LLC split).
  const units = useMemo<Unit[]>(() => rolls.flatMap((r) => r.units), [rolls]);

  const distinctSources = useMemo(
    () => [...new Set(units.map((u) => u.source).filter(Boolean))],
    [units],
  );
  const isGroup = Boolean(ownerSelected) || groupOverride || distinctSources.length > 1;

  const analysis = useMemo(
    () => (units.length ? analyze(units, { taxYear, utilityAllowance, forceGroup: Boolean(ownerSelected) || groupOverride, state: filingState }) : null),
    [units, taxYear, utilityAllowance, ownerSelected, groupOverride, filingState],
  );

  // Final config = derived hub facts + UI overrides.
  const config: CertConfig | null = useMemo(() => {
    if (!derived) return null;
    const c: CertConfig = JSON.parse(JSON.stringify(derived.config));
    c.certification.relationshipToOwner = relationship;
    c.property.description = description;
    c.property.taxMapParcels = [];
    c.filing.taxYear = taxYear;
    // For a group (or when the hub had no county), take counties from the rent rolls.
    const rollCounties = [...new Set(units.map((u) => u.county).filter((x): x is string => Boolean(x)))];
    if ((isGroup || c.property.counties.length === 0) && rollCounties.length) {
      c.property.counties = rollCounties;
      if (!c.property.addressLine) {
        c.property.addressLine = `Scattered sites located in ${rollCounties.join(' and ')} ${rollCounties.length > 1 ? 'Counties' : 'County'}, South Carolina`;
      }
    }
    // Jurisdiction (statute + addressee): default by state/counties, editable overrides.
    if (!representativeProperty) c.property.state = filingState; // prospect/paste use the selector
    const jd = jurisdictionDefaults(c.property.state, c.property.counties);
    c.jurisdiction = {
      statuteCitation: citationOverride.trim() || jd.statuteCitation,
      recipient: recipientOverride.trim() || jd.recipient,
    };
    // Nonprofit ownership = CAHP beneficial % through the chain, for this property.
    if (representativeProperty) {
      const i = cahpInterestFor(String(representativeProperty.id));
      c.nonprofit.ownershipPercent = i.ownershipPercent;
      c.nonprofit.memberClass = i.memberClass;
    }
    if (isGroup) {
      const gName = groupName || ownerSelected?.fields.Title || derived.config.company.legalName;
      if (ownerSelected?.fields.Title) c.company.legalName = ownerSelected.fields.Title;
      // Roster of the parent's subsidiary LLCs (name + EIN + nonprofit %) from the
      // hub. For a manual group (no parent owner), fall back to rent-roll source names.
      const members = ownerSelected
        ? childProperties.map((p) => {
            const i = cahpInterestFor(String(p.id));
            return { name: p.fields.Title || '', ein: p.fields.PropertyEIN || '', ownershipPercent: i.ownershipPercent, memberClass: i.memberClass };
          })
        : distinctSources.map((s) => ({ name: s, ein: '', ownershipPercent: null as number | null, memberClass: '' }));
      c.portfolio = {
        isGroupFiling: true,
        groupName: gName,
        groupStateType: 'South Carolina limited liability company',
        subsidiaryDescription: 'wholly-owned single-purpose subsidiary LLCs',
        members,
      };
    }
    return c;
  }, [derived, relationship, description, citationOverride, recipientOverride, taxYear, isGroup, groupName, ownerSelected, representativeProperty, filingState, units, distinctSources, childProperties, owners.data, ownership.data, cahpInterestFor]);

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
        downloadBlob(buildExhibitBlob(analysis, config), `${baseName}_Unit_AMI_Analysis_INTERNAL.xlsx`);
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
      const meta = { PropertyLookupId: String(representativeProperty.id) };
      // Save only the certification letter — the rent roll (not the unit analysis)
      // accompanies the submittal, so the Exhibit isn't uploaded here.
      await uploadDocument({
        libraryName: CERT_LIBRARY,
        filename: `${baseName}_Safe_Harbor_Certification_TY${taxYear}.docx`,
        file: docxBlob,
        metadata: meta,
      });
      setSaveMsg(
        `Saved the certification letter to "${CERT_LIBRARY}", tagged to ${config.company.legalName}. ` +
          `It now appears on that property's Documents tab.`,
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
  // Units that couldn't be classified and were defaulted to Market (informational).
  const nDefaulted = analysis
    ? analysis.units.filter((u) => !u.nonResidential && u.notes.some((n) => n.includes('counted as Market'))).length
    : 0;

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full p-5 my-8">
        <div className="flex items-start justify-between mb-1">
          <h3 className="text-lg font-bold text-teal-700">Generate Safe Harbor Certification</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600"><Icon name="x" size={18} /></button>
        </div>
        <p className="text-sm text-gray-600 mb-4">
          Upload AppFolio rent roll(s), pick the filing entity, and generate the §12-37-220(B)(11)(e) /
          Rev. Proc. 96-32 certification letter (.docx + PDF), plus an internal unit-AMI analysis. Qualification is rent-based;
          drop more than one rent roll to file a portfolio (group) of LLCs together.
        </p>

        {loading ? (
          <div className="py-8 text-center text-sm text-gray-500">Loading entities…</div>
        ) : (
          <div className="space-y-4">
            {/* Input-format toggle */}
            <div className="flex gap-1 text-xs">
              {(['appfolio', 'paste'] as const).map((m) => (
                <button
                  key={m}
                  onClick={() => setInputMode(m)}
                  className={`px-3 py-1 rounded border ${inputMode === m ? 'bg-teal-700 text-white border-teal-700' : 'border-gray-300 text-gray-600 hover:bg-gray-50'}`}
                >
                  {m === 'appfolio' ? 'AppFolio rent roll (.xlsx)' : 'Paste / PDF (other format)'}
                </button>
              ))}
            </div>

            {inputMode === 'appfolio' ? (
              /* Rent roll dropzone */
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
            ) : (
              /* Paste / PDF — non-AppFolio rent rolls from outside PMs */
              <div className="border border-gray-200 rounded-md p-3 space-y-2">
                <div className="text-xs text-gray-600">
                  Paste the rent-roll text, or upload a PDF to extract it. Each unit needs a line like
                  <span className="font-mono"> Unit 1 … $995</span>. Outside rent rolls rarely include bedrooms or
                  county — set the defaults below (applied to every parsed unit; county is auto-detected when possible).
                </div>
                <div className="grid grid-cols-4 gap-2">
                  <label className="flex flex-col gap-1 text-xs">
                    <span className="font-semibold text-gray-600">State</span>
                    <select value={manualState} onChange={(e) => { setManualState(e.target.value as 'SC' | 'NC'); setPasteCounty(''); }} className="border border-gray-300 rounded px-2 py-1">
                      <option value="SC">SC</option>
                      <option value="NC">NC</option>
                    </select>
                  </label>
                  <label className="flex flex-col gap-1 text-xs">
                    <span className="font-semibold text-gray-600">Property / source label</span>
                    <input value={pasteSource} onChange={(e) => setPasteSource(e.target.value)} placeholder="e.g. 700 Brook St" className="border border-gray-300 rounded px-2 py-1" />
                  </label>
                  <label className="flex flex-col gap-1 text-xs">
                    <span className="font-semibold text-gray-600">Default county</span>
                    <select value={pasteCounty} onChange={(e) => setPasteCounty(e.target.value)} className="border border-gray-300 rounded px-2 py-1">
                      <option value="">(auto-detect)</option>
                      {Object.keys(countiesForState(manualState)).sort().map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </label>
                  <label className="flex flex-col gap-1 text-xs">
                    <span className="font-semibold text-gray-600">Default bedrooms</span>
                    <input type="number" min={0} max={6} value={pasteBedrooms} onChange={(e) => setPasteBedrooms(e.target.value)} placeholder="e.g. 1" className="border border-gray-300 rounded px-2 py-1" />
                  </label>
                </div>
                <textarea
                  value={pasteText}
                  onChange={(e) => setPasteText(e.target.value)}
                  rows={6}
                  placeholder={'Paste rent-roll text here…\n\n700 Brook Street Unit 2  $995.00\n700 Brook Street Unit 8  $575.00\n…'}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs font-mono"
                />
                <div className="flex items-center gap-2">
                  <label className="px-3 py-1.5 border border-gray-300 rounded text-xs cursor-pointer hover:bg-gray-50">
                    {pdfBusy ? 'Reading PDF…' : 'Upload PDF'}
                    <input type="file" accept=".pdf" className="hidden" onChange={(e) => e.target.files?.[0] && handlePdfFile(e.target.files[0])} />
                  </label>
                  <button onClick={handleParseText} disabled={!pasteText.trim()} className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded text-xs font-medium disabled:opacity-50">Parse &amp; add units</button>
                </div>
              </div>
            )}
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

            {/* Prospect toggle — screen an entity that isn't in the hub yet */}
            <label className="flex items-center gap-2 text-xs text-gray-700 bg-amber-50 border border-amber-200 rounded px-2 py-1.5">
              <input type="checkbox" checked={prospectMode} onChange={(e) => setProspectMode(e.target.checked)} />
              <span><strong>Prospect / not in the system yet</strong> — qualification only; skips the entity lookup and Tax Map ID matching (use this to test a potential client before onboarding).</span>
            </label>

            {/* Entity + options */}
            <div className="grid grid-cols-2 gap-3 text-sm">
              {prospectMode ? (
                <>
                  <label className="flex flex-col gap-1">
                    <span className="text-xs font-semibold text-gray-600">Prospect / entity name</span>
                    <input value={prospectName} onChange={(e) => setProspectName(e.target.value)} placeholder="e.g. Prospect Holdings, LLC" className="border border-gray-300 rounded px-2 py-1" />
                  </label>
                  <label className="flex flex-col gap-1">
                    <span className="text-xs font-semibold text-gray-600">State</span>
                    <select value={manualState} onChange={(e) => setManualState(e.target.value as 'SC' | 'NC')} className="border border-gray-300 rounded px-2 py-1">
                      <option value="SC">SC</option>
                      <option value="NC">NC</option>
                    </select>
                  </label>
                </>
              ) : (
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
              )}
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
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-gray-600">Statute citation</span>
                <input value={citationOverride} onChange={(e) => setCitationOverride(e.target.value)} placeholder={config?.jurisdiction.statuteCitation || 'auto by state'} className="border border-gray-300 rounded px-2 py-1" />
              </label>
              <label className="flex flex-col gap-1">
                <span className="text-xs font-semibold text-gray-600">Addressed to</span>
                <input value={recipientOverride} onChange={(e) => setRecipientOverride(e.target.value)} placeholder={config?.jurisdiction.recipient || 'auto by state/county'} className="border border-gray-300 rounded px-2 py-1" />
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

            {/* Subsidiary roster preview (parent-owner group filing) */}
            {isGroup && ownerSelected && childProperties.length > 0 && (
              <div className="text-xs text-gray-600 border border-gray-200 rounded p-2">
                This parent filing will list <strong>{childProperties.length}</strong> subsidiary LLC(s) and
                their EINs on the document (from the hub).
              </div>
            )}

            {/* CAHP beneficial ownership readout (helps verify the nonprofit %) */}
            {ownerSelected && (
              <div className="text-xs bg-gray-50 border border-gray-200 rounded p-2 text-gray-700">
                CAHP beneficial ownership of <strong>{ownerSelected.fields.Title}</strong>:{' '}
                <strong>{parentCahpPct == null ? 'not found in the ownership chain' : `${parentCahpPct}%`}</strong>
                {parentCahpPct != null && ' — applied to each wholly-owned sub-LLC as its nonprofit ownership.'}
              </div>
            )}

            {/* Live determination preview */}
            {analysis && p && cnt && sc && (
              <div className="border border-gray-200 rounded-md p-3">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-sm font-bold text-gray-800">Determination</div>
                  <div className={`text-sm font-bold ${sc.qualifies.length ? 'text-success' : 'text-warning'}`}>{sc.headline}</div>
                </div>
                <div className="text-xs text-gray-600 mb-2">
                  {analysis.roll.denom} residential units{analysis.roll.nNonRes ? ` · ${analysis.roll.nNonRes} non-residential excluded` : ''}
                  {isGroup ? ` · ${distinctSources.length} LLCs` : ''}
                </div>
                <table className="w-full text-xs">
                  <thead className="text-[10px] text-gray-500 uppercase">
                    <tr><th className="text-left">AMI Tier</th><th>Units</th><th>%</th><th>Required</th><th>Result</th></tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const low: [string, number, number, string, boolean] = ['Low-Income (≤80%)', cnt.le80, p.le80, '≥75%', p.le80 >= 75];
                      const mkt: [string, number, number, string, boolean] = ['Market (>80%)', cnt.market, p.market, '≤25%', p.market <= 25];
                      const r50: [string, number, number, string, boolean] = ['Very Low (≤50%) — 50% AMI program', cnt.le50, p.le50, '≥20%', p.le50 >= 20];
                      const r60: [string, number, number, string, boolean] = ['≤60% — 60% AMI program', cnt.le60, p.le60, '≥40%', p.le60 >= 40];
                      return sc.chosen === '20/50' ? [r50, low, mkt]
                        : sc.chosen === '40/60' ? [r60, low, mkt]
                        : [low, r60, r50, mkt];
                    })().map(([label, u, pc, req, ok], i) => (
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
                {nDefaulted > 0 && (
                  <div className="mt-2 text-[11px] text-gray-500">
                    {nDefaulted} unit(s) were missing bedroom/rent/county and were counted as Market (the conservative default) — see the Notes column in the unit analysis.
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
            <button disabled={!ready || busy !== null} onClick={() => doDownload('xlsx')} className="px-3 py-1.5 border border-teal-300 text-teal-700 hover:bg-teal-50 rounded-md text-sm font-medium disabled:opacity-50">{busy === 'xlsx' ? '…' : 'Unit analysis .xlsx'}</button>
            <button disabled={!ready || !representativeProperty || prospectMode || busy !== null} title={prospectMode ? 'Prospect mode — download only (entity is not in the hub)' : !representativeProperty ? 'Pick a property (or an owner with at least one linked LLC) to save into the hub' : ''} onClick={saveToEntity} className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium disabled:opacity-50">{busy === 'save' ? 'Saving…' : 'Save to entity'}</button>
          </div>
        </div>
      </div>
    </div>
  );
}
