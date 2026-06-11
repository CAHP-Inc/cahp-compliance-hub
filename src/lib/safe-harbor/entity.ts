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
    ownershipPercent: number;
    memberClass: string;
    isTaxExempt: boolean;
  };
  /** How the certification is executed and who the signer is authorized for. */
  certification: {
    /** e.g. "property manager and authorized agent" — ties the signing company to the owner. */
    relationshipToOwner: string;
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
      parentEin: parent?.fields.TaxID || '',
      ownershipPercent: cahpRow?.fields.OwnershipPercent ?? DEFAULT_CERT.nonprofit.ownershipPercent,
      memberClass: cahpRow?.fields.MemberClass || DEFAULT_CERT.nonprofit.memberClass,
      isTaxExempt: true,
    },
    certification: { ...DEFAULT_CERT.certification },
    filing: {
      taxYear,
      filingType: 'Annual Renewal Certification',
      annualCertificationDeadline: 'October 1',
    },
  };

  return { config, exemptionChainOk, warnings };
}
