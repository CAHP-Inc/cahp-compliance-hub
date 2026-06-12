import { useCallback, useMemo, useState, useRef, type DragEvent } from 'react';
import JSZip from 'jszip';
import {
  useSharePointList,
  uploadDocument,
  getUpstreamOwnerIds,
  computeBeneficialOwnership,
  isCahpEntity,
  LIST_NAMES,
  type Owner,
  type Ownership,
  type Property,
  type TaxMapID,
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

// Normalize an entity name for matching a rent roll's "Property Groups" label to
// a hub LLC. Converts roman-numeral tokens to digits on BOTH sides (applied
// symmetrically, so the "IV" fund prefix is harmless): "IV SPB 2 LLC" and
// "IV SPB II LLC" both normalize to "4 spb 2 llc".
const ROMAN: Record<string, number> = {
  i: 1, ii: 2, iii: 3, iv: 4, v: 5, vi: 6, vii: 7, viii: 8, ix: 9, x: 10, xi: 11, xii: 12,
};
function normName(s: string | undefined): string {
  return (s || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .map((t) => (ROMAN[t] !== undefined ? String(ROMAN[t]) : t))
    .join(' ');
}
/** Best hub child-LLC id for a rent-roll source: exact normalized match, else max token overlap. */
function suggestChildId(source: string, children: { id: string; title: string }[]): string {
  const ns = normName(source);
  const exact = children.find((c) => normName(c.title) === ns);
  if (exact) return exact.id;
  const stoks = new Set(ns.split(' '));
  let best = '';
  let bestScore = 0;
  for (const c of children) {
    const overlap = normName(c.title).split(' ').filter((t) => stoks.has(t)).length;
    if (overlap > bestScore) {
      bestScore = overlap;
      best = c.id;
    }
  }
  return bestScore >= 2 ? best : '';
}

// ── Address matching (rent-roll unit -> Tax Map ID parcel -> hub LLC) ──
const STREET_SUFFIX: Record<string, string> = {
  street: 'st', avenue: 'ave', drive: 'dr', road: 'rd', circle: 'cir', lane: 'ln',
  court: 'ct', boulevard: 'blvd', place: 'pl', terrace: 'ter', parkway: 'pkwy',
  highway: 'hwy', trail: 'trl', cove: 'cv', square: 'sq', point: 'pt',
};
function normalizeAddr(raw: string | undefined): string {
  let s = (raw || '').toLowerCase();
  if (s.includes(' - ')) s = s.slice(s.lastIndexOf(' - ') + 3); // address after the display name
  s = s.split(',')[0]; // drop ", ST ZIP" tail
  s = s.replace(/\b\d{5}(-\d{4})?\b/g, ' '); // drop any stray zip
  s = s.replace(/[^a-z0-9]+/g, ' ').trim();
  return s.split(/\s+/).filter(Boolean).map((t) => STREET_SUFFIX[t] ?? t).join(' ');
}
interface ParcelEntry { num: string; street: string[]; propertyId: string; title: string }
function buildParcelIndex(
  taxmaps: { fields: { ParcelAddress?: string; LinkedPropertyLookupId?: string } }[] | null | undefined,
  titleById: Map<string, string>,
): ParcelEntry[] {
  const out: ParcelEntry[] = [];
  for (const t of taxmaps ?? []) {
    const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
    if (!t.fields.ParcelAddress || !pid) continue;
    const toks = normalizeAddr(t.fields.ParcelAddress).split(' ').filter(Boolean);
    if (!toks.length || !/^\d/.test(toks[0])) continue;
    out.push({ num: toks[0], street: toks.slice(1), propertyId: pid, title: titleById.get(pid) || '' });
  }
  return out;
}
/** Match a rent-roll unit address to the best parcel (same house number + street overlap). */
function matchParcel(rawUnitAddr: string, index: ParcelEntry[]): ParcelEntry | null {
  const toks = normalizeAddr(rawUnitAddr).split(' ').filter(Boolean);
  if (!toks.length || !/^\d/.test(toks[0])) return null;
  const num = toks[0];
  const street = toks.slice(1);
  let best: ParcelEntry | null = null;
  let bestScore = 0;
  for (const p of index) {
    if (p.num !== num) continue;
    const overlap = p.street.filter((t) => street.includes(t)).length;
    if (overlap > bestScore) {
      bestScore = overlap;
      best = p;
    }
  }
  return bestScore > 0 ? best : null;
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
  const taxmaps = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 500 });
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
  // Per-rent-roll override: roll index -> hub child Property id ('' = none/keep label).
  const [rollMapOverride, setRollMapOverride] = useState<Record<number, string>>({});
  // Manual per-unit assignment for units with no Tax Map ID match: "ri-ui" -> property id ('' = exclude).
  const [unitAssign, setUnitAssign] = useState<Record<string, string>>({});
  // Prospect mode: entity isn't in the hub yet — qualification only, skip TMID/entity.
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
    if (prospectMode) {
      return { config: defaultCertConfig(prospectName || 'Prospective Entity', taxYear), exemptionChainOk: true, warnings: [] };
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
      return { config: defaultCertConfig(ownerSelected.fields.Title ?? '', taxYear), exemptionChainOk: true, warnings: [] };
    }
    return null;
  }, [prospectMode, prospectName, representativeProperty, ownerSelected, owners.data, ownership.data, taxYear]);

  // Hub child LLCs available to map rent rolls onto (only when filing for an owner).
  const childOptions = useMemo(
    () => childProperties.map((p) => ({ id: String(p.id), title: p.fields.Title || '' })),
    [childProperties],
  );
  const useHubNames = !prospectMode && selType === 'owner' && childOptions.length > 0;

  // Effective rent-roll -> hub-LLC mapping (auto-suggested, user-overridable).
  const rollMap = useMemo(
    () =>
      rolls.map((r, i) =>
        rollMapOverride[i] !== undefined ? rollMapOverride[i] : suggestChildId(r.source, childOptions),
      ),
    [rolls, rollMapOverride, childOptions],
  );

  // Parcel index (Tax Map ID ParcelAddress -> hub LLC), used to split a single
  // combined rent roll across the right subsidiary LLCs by unit address.
  const propTitleById = useMemo(
    () => new Map((properties.data ?? []).map((p) => [String(p.id), p.fields.Title || ''])),
    [properties.data],
  );
  const parcelIndex = useMemo(() => buildParcelIndex(taxmaps.data, propTitleById), [taxmaps.data, propTitleById]);
  const useParcelMatch = !prospectMode && selType === 'owner' && parcelIndex.length > 0;

  // CAHP's beneficial ownership % of the SELECTED PARENT fund (computed once).
  // Its wholly-owned subsidiary LLCs inherit this %, so per-LLC certs are correct
  // even before each sub-LLC's own ownership row is entered in the hub.
  const parentCahpPct = useMemo(() => {
    if (!ownerSelected || !owners.data || !ownership.data) return null;
    const bens = computeBeneficialOwnership('owner', String(ownerSelected.id), ownership.data, owners.data);
    const cahp = bens.find((b) => isCahpEntity(b.owner));
    return cahp ? cahp.beneficialPercent : null;
  }, [ownerSelected, owners.data, ownership.data]);

  // CAHP nonprofit's beneficial ownership %/class for a property. For a parent-fund
  // filing it flows from the parent (× the parent's stake in the sub, default 100%);
  // otherwise it's compounded up the property's own ownership chain. Member class
  // comes from a direct CAHP member row when present (e.g. a "Class C" structure).
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
          )?.fields.OwnershipPercent ?? 100; // default: wholly-owned single-purpose LLC
        return { ownershipPercent: round((parentCahpPct * parentOwnPct) / 100), memberClass };
      }

      const bens = computeBeneficialOwnership('property', propertyId, ownership.data, owners.data);
      const cahp = bens.find((b) => isCahpEntity(b.owner));
      return { ownershipPercent: cahp ? round(cahp.beneficialPercent) : null, memberClass };
    },
    [owners.data, ownership.data, ownerSelected, parentCahpPct],
  );

  // Tax Map ID parcel numbers registered to a property (for the per-LLC filing).
  const parcelsForProperty = useCallback(
    (propertyId: string): string[] =>
      (taxmaps.data ?? [])
        .filter((t) => String(t.fields.LinkedPropertyLookupId) === propertyId && t.fields.Title)
        .map((t) => String(t.fields.Title)),
    [taxmaps.data],
  );
  // LLC options for manually assigning an unmatched unit.
  const assignOptions = useMemo(
    () =>
      (properties.data ?? [])
        .map((p) => ({ id: String(p.id), title: p.fields.Title || '' }))
        .sort((a, b) => a.title.localeCompare(b.title)),
    [properties.data],
  );

  // Units fed to the analysis. When filing for an owner with parcels, each unit
  // is tied to its subsidiary LLC by Tax Map ID address match; unmatched units
  // can be assigned manually, and any still-unmatched unit is EXCLUDED from the
  // certification (it doesn't correlate to a hub parcel). Otherwise units keep
  // the per-rent-roll mapping or their "Property Groups" label.
  const proc = useMemo(() => {
    if (!useHubNames && !useParcelMatch) {
      return { units: rolls.flatMap((r) => r.units), autoMatched: 0, manualAssigned: 0, excluded: 0, nonMatched: [] as { key: string; u: Unit; assignId: string }[] };
    }
    let autoMatched = 0;
    let manualAssigned = 0;
    let excluded = 0;
    const included: Unit[] = [];
    const nonMatched: { key: string; u: Unit; assignId: string }[] = [];
    rolls.forEach((r, ri) => {
      const rollHubName = useHubNames ? childOptions.find((c) => c.id === rollMap[ri])?.title : undefined;
      r.units.forEach((u, ui) => {
        if (!useParcelMatch) {
          included.push({ ...u, source: rollHubName || u.source, notes: [] });
          return;
        }
        const key = `${ri}-${ui}`;
        const auto = matchParcel(u.prop, parcelIndex);
        const autoId = auto?.propertyId ?? '';
        if (autoId) autoMatched++;
        const assignId = unitAssign[key] !== undefined ? unitAssign[key] : autoId;
        if (!autoId) {
          nonMatched.push({ key, u, assignId });
          if (assignId) manualAssigned++;
          else excluded++;
        }
        if (!assignId) return; // not correlated to a parcel -> excluded
        included.push({ ...u, source: propTitleById.get(assignId) || u.source, notes: [] });
      });
    });
    return { units: included, autoMatched, manualAssigned, excluded, nonMatched };
  }, [rolls, useHubNames, useParcelMatch, childOptions, rollMap, parcelIndex, unitAssign, propTitleById]);
  const units = proc.units;

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
    // Nonprofit ownership = CAHP beneficial % through the chain, for this property.
    if (representativeProperty) {
      const i = cahpInterestFor(String(representativeProperty.id));
      c.nonprofit.ownershipPercent = i.ownershipPercent;
      c.nonprofit.memberClass = i.memberClass;
    }
    if (isGroup) {
      const gName = groupName || ownerSelected?.fields.Title || derived.config.company.legalName;
      if (ownerSelected?.fields.Title) c.company.legalName = ownerSelected.fields.Title;
      // Per-subsidiary nonprofit beneficial ownership from the hub (varies by LLC).
      const propByTitle = new Map((properties.data ?? []).map((p) => [p.fields.Title || '', p]));
      const members = distinctSources.map((s) => {
        const prop = propByTitle.get(s);
        const own = prop ? cahpInterestFor(String(prop.id)) : { ownershipPercent: null, memberClass: '' };
        return { name: s, ownershipPercent: own.ownershipPercent, memberClass: own.memberClass };
      });
      c.portfolio = {
        isGroupFiling: true,
        groupName: gName,
        groupStateType: 'South Carolina limited liability company',
        subsidiaryDescription: 'wholly-owned single-purpose subsidiary LLCs',
        members,
      };
    }
    return c;
  }, [derived, relationship, description, parcels, taxYear, isGroup, groupName, ownerSelected, representativeProperty, units, distinctSources, properties.data, owners.data, ownership.data, cahpInterestFor]);

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

  // Build a single-LLC config for a sub-entity under the parent (its own EIN,
  // counties, Tax Map IDs, and CAHP beneficial ownership).
  const buildLlcConfig = (property: Property, llcUnits: Unit[]): CertConfig => {
    const base: CertConfig = derived
      ? JSON.parse(JSON.stringify(derived.config))
      : defaultCertConfig(property.fields.Title || '', taxYear, property.fields.cahpState || 'SC');
    base.certification.relationshipToOwner = relationship;
    base.filing.taxYear = taxYear;
    base.company = {
      legalName: property.fields.Title || '',
      stateType: 'South Carolina limited liability company',
      ein: property.fields.PropertyEIN || '',
      dorAccountId: property.fields.DORAccountID || '',
    };
    const counties = [...new Set(llcUnits.map((u) => u.county).filter((x): x is string => Boolean(x)))];
    base.property = {
      description,
      addressLine: property.fields.PropertyAddress
        || (counties.length ? `Scattered sites located in ${counties.join(' and ')} ${counties.length > 1 ? 'Counties' : 'County'}, South Carolina` : ''),
      counties,
      state: property.fields.cahpState || 'SC',
      taxMapParcels: parcelsForProperty(String(property.id)),
    };
    const interest = cahpInterestFor(String(property.id));
    base.nonprofit.ownershipPercent = interest.ownershipPercent;
    base.nonprofit.memberClass = interest.memberClass;
    delete base.portfolio; // each sub-LLC files as its own single entity
    return base;
  };

  // Generate one certification per sub-LLC under the parent, zipped.
  const generatePerEntity = async () => {
    if (!properties.data) return;
    setBusy('perentity');
    setSaveErr(null);
    setSaveMsg(null);
    try {
      const propByTitle = new Map(properties.data.map((p) => [p.fields.Title || '', p]));
      const zip = new JSZip();
      let count = 0;
      const skipped: string[] = [];
      for (const s of distinctSources) {
        const prop = propByTitle.get(s);
        const llcUnits = units.filter((u) => u.source === s).map((u) => ({ ...u, notes: [] as string[] }));
        if (!prop || !llcUnits.length) {
          if (llcUnits.length) skipped.push(s);
          continue;
        }
        const cfg = buildLlcConfig(prop, llcUnits);
        const a = analyze(llcUnits, { taxYear, utilityAllowance, forceGroup: false });
        const slug = slugify(cfg.company.legalName);
        zip.file(`${slug}_Safe_Harbor_Certification_TY${taxYear}.docx`, await buildLetterDocx(a, cfg));
        zip.file(`${slug}_Safe_Harbor_Certification_TY${taxYear}.pdf`, buildLetterPdf(a, cfg));
        zip.file(`${slug}_Unit_AMI_Analysis_INTERNAL.xlsx`, buildExhibitBlob(a, cfg));
        count++;
      }
      if (!count) {
        setSaveErr('No sub-LLCs with units matched to hub entities. Match units to LLCs (by Tax Map ID or manually) first.');
        return;
      }
      const blob = await zip.generateAsync({ type: 'blob' });
      const parentSlug = slugify(ownerSelected?.fields.Title || prospectName || 'Portfolio');
      downloadBlob(blob, `${parentSlug}_Per_Entity_Safe_Harbor_TY${taxYear}.zip`);
      setSaveMsg(
        `Generated ${count} per-entity certification(s) (.docx + PDF + internal analysis each) in the zip.` +
          (skipped.length ? ` Skipped ${skipped.length} group(s) with no matching hub LLC: ${skipped.join(', ')}.` : ''),
      );
    } catch (e) {
      setSaveErr(e instanceof Error ? e.message : String(e));
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

            {/* Prospect toggle — screen an entity that isn't in the hub yet */}
            <label className="flex items-center gap-2 text-xs text-gray-700 bg-amber-50 border border-amber-200 rounded px-2 py-1.5">
              <input type="checkbox" checked={prospectMode} onChange={(e) => setProspectMode(e.target.checked)} />
              <span><strong>Prospect / not in the system yet</strong> — qualification only; skips the entity lookup and Tax Map ID matching (use this to test a potential client before onboarding).</span>
            </label>

            {/* Entity + options */}
            <div className="grid grid-cols-2 gap-3 text-sm">
              {prospectMode ? (
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-semibold text-gray-600">Prospect / entity name</span>
                  <input value={prospectName} onChange={(e) => setProspectName(e.target.value)} placeholder="e.g. Prospect Holdings, LLC" className="border border-gray-300 rounded px-2 py-1" />
                </label>
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

            {/* Parcel-address (Tax Map ID) match summary + manual assignment */}
            {useParcelMatch && rolls.length > 0 && (
              <div className="border border-teal-200 rounded-md">
                <div className="bg-teal-50 rounded-t-md p-2 text-xs text-teal-900">
                  <strong>{proc.autoMatched}</strong> matched to LLCs by parcel address ·{' '}
                  <strong>{proc.manualAssigned}</strong> assigned manually ·{' '}
                  <strong>{proc.excluded}</strong> excluded (no Tax Map ID match).
                  {proc.excluded > 0 && ' Excluded units are NOT in the certification.'}
                </div>
                {proc.nonMatched.length > 0 && (
                  <div className="p-2">
                    <div className="text-[11px] text-gray-600 mb-1">
                      Units with no Tax Map ID match — assign an LLC to include them, or leave “Exclude”:
                    </div>
                    <div className="space-y-1 max-h-44 overflow-y-auto">
                      {proc.nonMatched.map(({ key, u }) => (
                        <div key={key} className="flex items-center gap-2 text-xs">
                          <span className="text-gray-500 flex-1 truncate" title={u.prop}>
                            {u.prop}{u.unit ? ` [${u.unit}]` : ''}
                          </span>
                          <span className="text-gray-300">→</span>
                          <select
                            value={unitAssign[key] ?? ''}
                            onChange={(e) => setUnitAssign((prev) => ({ ...prev, [key]: e.target.value }))}
                            className="border border-gray-300 rounded px-2 py-1 w-56"
                          >
                            <option value="">— Exclude (no parcel) —</option>
                            {assignOptions.map((c) => (
                              <option key={c.id} value={c.id}>{c.title}</option>
                            ))}
                          </select>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Map each rent roll to the parent owner's subsidiary LLC in the hub */}
            {useHubNames && !useParcelMatch && rolls.length > 0 && (
              <div className="border border-gray-200 rounded-md p-3">
                <div className="text-xs font-semibold text-gray-700 mb-2">
                  Map rent rolls to the hub's subsidiary LLCs
                  <span className="font-normal text-gray-500"> — the document lists these registered names, not the rent-roll labels.</span>
                </div>
                <div className="space-y-1">
                  {rolls.map((r, i) => (
                    <div key={i} className="flex items-center gap-2 text-xs">
                      <span className="text-gray-500 w-40 truncate" title={r.source}>{r.source}</span>
                      <span className="text-gray-300">→</span>
                      <select
                        value={rollMap[i]}
                        onChange={(e) => setRollMapOverride((prev) => ({ ...prev, [i]: e.target.value }))}
                        className="border border-gray-300 rounded px-2 py-1 flex-1"
                      >
                        <option value="">— keep rent-roll label “{r.source}” —</option>
                        {childOptions.map((c) => (
                          <option key={c.id} value={c.id}>{c.title}</option>
                        ))}
                      </select>
                    </div>
                  ))}
                </div>
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
            {selType === 'owner' && !prospectMode && (
              <button disabled={!ready || busy !== null} title="One certification per sub-LLC (its own EIN + Tax Map IDs), bundled as a zip" onClick={generatePerEntity} className="px-3 py-1.5 border border-gold-500 bg-gold-50 text-gold-700 hover:bg-gold-200 rounded-md text-sm font-medium disabled:opacity-50">{busy === 'perentity' ? 'Zipping…' : 'Per-entity certs (.zip)'}</button>
            )}
            <button disabled={!ready || !representativeProperty || prospectMode || busy !== null} title={prospectMode ? 'Prospect mode — download only (entity is not in the hub)' : !representativeProperty ? 'Pick a property (or an owner with at least one linked LLC) to save into the hub' : ''} onClick={saveToEntity} className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium disabled:opacity-50">{busy === 'save' ? 'Saving…' : 'Save to entity'}</button>
          </div>
        </div>
      </div>
    </div>
  );
}
