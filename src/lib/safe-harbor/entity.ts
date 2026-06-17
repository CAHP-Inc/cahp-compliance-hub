/**
 * Safe Harbor certification config — the entity/boilerplate facts the letter
 * needs, plus a pure derivation from the hub's Owners / Ownership / Properties
 * data so the modal doesn't hand-type anything that already lives in the hub.
 *
 * The certifying party is the PROPERTY MANAGEMENT COMPANY that manages the
 * property (not the nonprofit). The nonprofit ownership chain is still described
 * (it qualifies the exemption), but the certification is executed by the PM
 * company as authorized agent for the owner(s). The signature block itself is
 * left fully blank and filled at signing.
 */

export interface CertConfig {
  company: {
    legalName: string;
    stateType: string;
    ein: string;
    dorAccountId: string;
  };
  property: {
    description: string;
    addressLine: string;
    counties: string[];
    state: string;
    taxMapParcels: string[];
  };
  nonprofit: {
    managingMemberName: string;
    parentName: string;
    parentEin: string;
    /** From the hub's Ownership Structure row. null = unknown (renders blank). */
    ownershipPercent: number | null;
    memberClass: string;
    isTaxExempt: boolean;
  };
  /** How the certification is executed and who the signer is authorized for. */
  certification: {
    /** e.g. "property manager and authorized agent" — ties the signing company to the owner. */
    relationshipToOwner: string;
  };
  /** State-specific legal framing (SC vs NC). Editable in the modal to adjust. */
  jurisdiction: {
    statuteCitation: string;   // e.g. "South Carolina Code §12-37-220(B)(11)(e)" or "N.C.G.S. §105-278.6"
    recipient: string;         // addressee, e.g. "the ... Department of Revenue and the X County Assessor"
  };
  filing: {
    taxYear: number;
    filingType: string;
    annualCertificationDeadline: string;
  };
  portfolio?: {
    isGroupFiling: boolean;
    groupName: string;
    groupStateType: string;
    subsidiaryDescription: string;
    /** Per-subsidiary nonprofit ownership from the hub (varies by LLC). Keyed by LLC name. */
    members?: { name: string; ownershipPercent: number | null; memberClass: string }[];
  };
}

/** Known CAHP 501(c)(3) parent EINs by state (public nonprofit identifiers). */
export const CAHP_EIN_BY_STATE: Record<string, string> = {
  SC: '99-4885069',
};

/**
 * State-specific statute citation + addressee for the certification.
 * NC NOTE: NC currently exempts nonprofit low-income housing under a charitable-use
 * test (N.C.G.S. §105-278.6), filed on Form AV-10 with the COUNTY tax assessor —
 * NOT the Rev. Proc. 96-32 safe harbor (a pending bill, H1042 / §105-278.7A, would
 * change that). These defaults are editable in the modal; confirm with NC counsel.
 */
export function jurisdictionDefaults(state: string, counties: string[]): { statuteCitation: string; recipient: string } {
  const countiesStr = counties.length
    ? counties.join(' and ') + ' ' + (counties.length > 1 ? 'Counties' : 'County')
    : 'County';
  if ((state || '').toUpperCase() === 'NC') {
    return { statuteCitation: 'N.C.G.S. §105-278.6', recipient: `the ${countiesStr} Tax Assessor` };
  }
  return {
    statuteCitation: 'South Carolina Code §12-37-220(B)(11)(e)',
    recipient: `the South Carolina Department of Revenue and the ${countiesStr} Assessor`,
  };
}

export const DEFAULT_CERT: Pick<CertConfig, 'certification'> & {
  nonprofit: Pick<CertConfig['nonprofit'], 'managingMemberName' | 'parentName' | 'memberClass' | 'ownershipPercent'>;
} = {
  certification: { relationshipToOwner: 'property manager and authorized agent' },
  nonprofit: {
    managingMemberName: 'CAHP SC, LLC',
    parentName: 'Carolina Affordable Housing Project Inc.',
    memberClass: 'Class C',
    ownershipPercent: 1,
  },
};

// Minimal shapes we read off the hub list items (kept local so this file has no
// dependency on the full SharePoint types — the modal passes plain objects).
export interface PropertyLike {
  id: string;
  fields: {
    Title: string;
    LegalEntity?: string;
    PropertyAddress?: string;
    UnitCount?: number;
    cahpCounty?: string;
    cahpState?: string;
    DORAccountID?: string;
    PropertyEIN?: string;
  };
}
export interface OwnerLike {
  id: string;
  fields: {
    Title?: string;
    OwnerType?: string;
    TaxID?: string;
    IsCAHPEntity?: boolean;
    IsTaxExempt?: boolean;
  };
}
export interface OwnershipLike {
  id: string;
  fields: {
    OwnerLookupId?: string | number;
    LinkedPropertyLookupId?: string | number;
    OwnershipPercent?: number;
    MemberClass?: string;
    RelationshipType?: string;
  };
}

export interface DerivedConfig {
  config: CertConfig;
  exemptionChainOk: boolean;
  warnings: string[];
}

/** Minimal config for a portfolio/owner filing when no representative child
 *  property is available to derive from. CAHP/EIN fields fall back to defaults. */
export function defaultCertConfig(legalName: string, taxYear: number, state = 'SC'): CertConfig {
  return {
    company: { legalName, stateType: 'South Carolina limited liability company', ein: '', dorAccountId: '' },
    property: { description: 'scattered-site residential rental units', addressLine: '', counties: [], state, taxMapParcels: [] },
    nonprofit: {
      managingMemberName: DEFAULT_CERT.nonprofit.managingMemberName,
      parentName: DEFAULT_CERT.nonprofit.parentName,
      parentEin: CAHP_EIN_BY_STATE[state] || '',
      ownershipPercent: null, // unknown until set in the hub
      memberClass: '',
      isTaxExempt: true,
    },
    certification: { ...DEFAULT_CERT.certification },
    jurisdiction: jurisdictionDefaults(state, []),
    filing: { taxYear, filingType: 'Annual Renewal Certification', annualCertificationDeadline: 'October 1' },
  };
}

/** The CAHP nonprofit's ownership %/class for one property, from the hub. */
export function cahpOwnershipForProperty(
  propertyId: string,
  owners: OwnerLike[],
  ownership: OwnershipLike[],
): { ownershipPercent: number | null; memberClass: string } {
  const ownerById = new Map(owners.map((o) => [String(o.id), o]));
  const row = ownership.find(
    (r) =>
      String(r.fields.LinkedPropertyLookupId) === String(propertyId) &&
      ownerById.get(String(r.fields.OwnerLookupId))?.fields.IsCAHPEntity,
  );
  return { ownershipPercent: row?.fields.OwnershipPercent ?? null, memberClass: row?.fields.MemberClass || '' };
}

function parseCounties(value: string | undefined): string[] {
  if (!value) return [];
  return value
    .split(',')
    .map((s) => s.replace(/\((SC|NC)\)/g, '').trim())
    .filter(Boolean);
}

/**
 * Derive a CertConfig for one property from hub data.
 *
 * @param upstreamOwnerIds set of owner IDs upstream of this property (from
 *   getUpstreamOwnerIds) — used to confirm a CAHP 501(c)(3) is in the chain.
 */
export function deriveCertConfig(args: {
  property: PropertyLike;
  owners: OwnerLike[];
  ownership: OwnershipLike[];
  upstreamOwnerIds: Set<string>;
  taxYear: number;
}): DerivedConfig {
  const { property, owners, ownership, upstreamOwnerIds, taxYear } = args;
  const warnings: string[] = [];
  const ownerById = new Map(owners.map((o) => [String(o.id), o]));

  const directRows = ownership.filter(
    (r) => String(r.fields.LinkedPropertyLookupId) === String(property.id),
  );

  // CAHP instrumentality member among this property's direct owners.
  const cahpRow = directRows.find((r) => {
    const o = ownerById.get(String(r.fields.OwnerLookupId));
    return o?.fields.IsCAHPEntity;
  });
  const cahpMember = cahpRow ? ownerById.get(String(cahpRow.fields.OwnerLookupId)) : undefined;

  // Any CAHP entity upstream counts toward the exemption chain.
  const upstreamHasCahp = [...upstreamOwnerIds].some(
    (id) => ownerById.get(id)?.fields.IsCAHPEntity,
  );
  const exemptionChainOk = Boolean(cahpMember) || upstreamHasCahp;
  if (!exemptionChainOk) {
    warnings.push(
      `No CAHP 501(c)(3) instrumentality found in the ownership of "${property.fields.Title}". ` +
        'Without that nonprofit member the entity does not qualify for the §12-37-220(B)(11)(e) ' +
        'exemption — verify the ownership before filing.',
    );
  }

  // Parent nonprofit corp (holds the 501(c)(3) EIN).
  const parent = owners.find(
    (o) => o.fields.IsCAHPEntity && (o.fields.OwnerType === 'Nonprofit' || o.fields.IsTaxExempt),
  );

  const counties = parseCounties(property.fields.cahpCounty);

  const config: CertConfig = {
    company: {
      legalName: property.fields.Title || property.fields.LegalEntity || '',
      stateType: 'South Carolina limited liability company',
      ein: property.fields.PropertyEIN || '',
      dorAccountId: property.fields.DORAccountID || '',
    },
    property: {
      description: 'scattered-site residential rental units',
      addressLine: property.fields.PropertyAddress || '',
      counties,
      state: property.fields.cahpState || 'SC',
      taxMapParcels: [],
    },
    nonprofit: {
      managingMemberName: cahpMember?.fields.Title || DEFAULT_CERT.nonprofit.managingMemberName,
      parentName: parent?.fields.Title || DEFAULT_CERT.nonprofit.parentName,
      // EIN: hub's CAHP parent TaxID, else the known state CAHP EIN.
      parentEin: parent?.fields.TaxID || CAHP_EIN_BY_STATE[property.fields.cahpState || 'SC'] || '',
      // Ownership %/class come from THIS LLC's hub ownership row — not a uniform
      // default (it varies by LLC). Blank when the hub doesn't have it.
      ownershipPercent: cahpRow?.fields.OwnershipPercent ?? null,
      memberClass: cahpRow?.fields.MemberClass || '',
      isTaxExempt: true,
    },
    certification: { ...DEFAULT_CERT.certification },
    jurisdiction: jurisdictionDefaults(property.fields.cahpState || 'SC', counties),
    filing: {
      taxYear,
      filingType: 'Annual Renewal Certification',
      annualCertificationDeadline: 'October 1',
    },
  };

  return { config, exemptionChainOk, warnings };
}
