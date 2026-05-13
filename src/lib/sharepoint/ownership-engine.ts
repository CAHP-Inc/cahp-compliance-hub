/**
 * Recursive ownership engine.
 *
 * Spec reference: Platform Specification §5 (Core Engine: Recursive Ownership Computation).
 *
 * Given a subject (property or owner entity), walks the Ownership Structure list to compute:
 *   - direct owners: immediate members of the subject
 *   - beneficial ownership tree: full chain to natural-person ultimate owners
 *
 * The engine operates on data already fetched via the standard hooks (useSharePointList),
 * so consumers must provide the full ownership and owners arrays. This keeps the engine
 * pure and testable.
 */

import type { Owner, Ownership } from './types';

export type SubjectType = 'property' | 'owner';

export interface OwnershipNode {
  /** The ownership relationship record */
  relationship: Ownership;
  /** The Owner entity referenced by this relationship */
  owner: Owner | null;
  /** Children (members of this owner, recursive) */
  children: OwnershipNode[];
  /** Depth from root (0 = direct owner of the original subject) */
  depth: number;
}

/**
 * Returns the direct ownership relationships for a given subject.
 *
 * @param subjectType 'property' or 'owner'
 * @param subjectId SharePoint integer ID (as string) of the property or owner
 * @param allOwnership Full Ownership Structure list (call useSharePointList ahead of this)
 */
export function getDirectOwnerships(
  subjectType: SubjectType,
  subjectId: string,
  allOwnership: Ownership[]
): Ownership[] {
  if (subjectType === 'property') {
    return allOwnership.filter(
      (o) => String(o.fields.LinkedPropertyLookupId) === String(subjectId)
    );
  }
  // subject is another owner — find rows where ParentOwner points to this entity
  return allOwnership.filter(
    (o) => String(o.fields.ParentOwnerLookupId) === String(subjectId)
  );
}

/**
 * Returns the direct owner ENTITIES (resolved from the relationships) for a subject,
 * sorted by ownership percentage descending.
 */
export function getDirectOwnersOf(
  subjectType: SubjectType,
  subjectId: string,
  allOwnership: Ownership[],
  allOwners: Owner[]
): { relationship: Ownership; owner: Owner | null }[] {
  const ownerById = new Map(allOwners.map((o) => [String(o.id), o]));
  return getDirectOwnerships(subjectType, subjectId, allOwnership)
    .map((rel) => ({
      relationship: rel,
      owner: rel.fields.OwnerLookupId
        ? ownerById.get(String(rel.fields.OwnerLookupId)) ?? null
        : null,
    }))
    .sort(
      (a, b) =>
        (b.relationship.fields.OwnershipPercent ?? 0) -
        (a.relationship.fields.OwnershipPercent ?? 0)
    );
}

/**
 * Walks the ownership chain recursively from a subject to natural-person owners.
 * Cycle-safe: tracks visited owner IDs to prevent infinite recursion.
 *
 * @returns root node whose children are direct owners of the subject; each child node
 *          itself has children if its owner has upstream members.
 */
export function getBeneficialOwnershipTree(
  subjectType: SubjectType,
  subjectId: string,
  allOwnership: Ownership[],
  allOwners: Owner[],
  maxDepth = 10
): OwnershipNode[] {
  const visited = new Set<string>();

  function walk(
    subjType: SubjectType,
    subjId: string,
    depth: number
  ): OwnershipNode[] {
    if (depth > maxDepth) return [];

    const direct = getDirectOwnersOf(subjType, subjId, allOwnership, allOwners);

    return direct.map((d) => {
      const ownerId = d.owner ? String(d.owner.id) : null;
      let children: OwnershipNode[] = [];

      if (ownerId && !visited.has(ownerId)) {
        // Only LLCs and Nonprofits can have upstream members; Individuals are leaves
        const ownerType = d.owner?.fields.OwnerType;
        if (ownerType === 'LLC' || ownerType === 'Nonprofit') {
          visited.add(ownerId);
          children = walk('owner', ownerId, depth + 1);
          visited.delete(ownerId);
        }
      }

      return {
        relationship: d.relationship,
        owner: d.owner,
        children,
        depth,
      };
    });
  }

  return walk(subjectType, subjectId, 0);
}

/**
 * Counts the number of properties an owner has direct or indirect interest in.
 * Direct = the owner appears in ownership rows linked to properties.
 * Indirect = the owner is upstream of another owner that has direct interest in properties.
 */
export function countPropertiesForOwner(
  ownerId: string,
  allOwnership: Ownership[],
  _allOwners: Owner[]
): { direct: number; indirect: number } {
  // Direct: rows where Owner == ownerId AND LinkedProperty is set
  const direct = allOwnership.filter(
    (o) =>
      String(o.fields.OwnerLookupId) === String(ownerId) &&
      o.fields.LinkedPropertyLookupId
  ).length;

  // Indirect: recursively walk DOWN from this owner — find owners that this owner has interest in,
  // then find their property holdings, repeat.
  const visited = new Set<string>([String(ownerId)]);
  const propertiesFound = new Set<string>();

  function walkDown(currentOwnerId: string) {
    // Find rows where ParentOwner == currentOwnerId — these are entities this owner has interest in
    const childRelations = allOwnership.filter(
      (o) => String(o.fields.ParentOwnerLookupId) === String(currentOwnerId)
    );
    for (const rel of childRelations) {
      const childOwnerId = rel.fields.OwnerLookupId;
      if (!childOwnerId || visited.has(String(childOwnerId))) continue;
      visited.add(String(childOwnerId));

      // Properties held by this child owner
      const childProps = allOwnership.filter(
        (o) =>
          String(o.fields.OwnerLookupId) === String(childOwnerId) &&
          o.fields.LinkedPropertyLookupId
      );
      for (const prop of childProps) {
        if (prop.fields.LinkedPropertyLookupId) {
          propertiesFound.add(String(prop.fields.LinkedPropertyLookupId));
        }
      }

      walkDown(String(childOwnerId));
    }
  }

  walkDown(String(ownerId));

  return { direct, indirect: propertiesFound.size };
}

/**
 * Counts the LLCs (or other entities) that an owner has direct interest in.
 * Used on Owner Detail to show "owns interest in N LLCs."
 */
export function countLLCsOwnedBy(ownerId: string, allOwnership: Ownership[]): number {
  return allOwnership.filter(
    (o) =>
      String(o.fields.OwnerLookupId) === String(ownerId) &&
      o.fields.ParentOwnerLookupId &&
      !o.fields.LinkedPropertyLookupId
  ).length;
}

/**
 * Counts the members of an entity (LLC or Nonprofit) — the owners that have direct interest in this entity.
 */
export function countMembersOf(ownerId: string, allOwnership: Ownership[]): number {
  return allOwnership.filter(
    (o) => String(o.fields.ParentOwnerLookupId) === String(ownerId)
  ).length;
}

// =============================================================================
// Beneficial ownership computation
// =============================================================================

export interface BeneficialOwner {
  /** The natural person or nonprofit at the top of an ownership chain */
  owner: Owner;
  /** Effective beneficial ownership of the subject (compounded percentages) */
  beneficialPercent: number;
  /** Each path through the chain that contributes to this beneficial owner's stake */
  paths: BeneficialPath[];
}

export interface BeneficialPath {
  /** The intermediate entities walked, from subject up to terminal owner (exclusive of terminal) */
  intermediates: { owner: Owner; relationship: Ownership }[];
  /** The terminal relationship (the row linking the natural person / nonprofit to the intermediate) */
  terminalRelationship: Ownership;
  /** The percentage of the subject this path represents */
  pathPercent: number;
}

/**
 * Walks the ownership tree from a subject, multiplying percentages along the way,
 * to produce a list of beneficial owners — the natural persons and nonprofits at the
 * top of the chain who hold the actual economic interest in the subject.
 *
 * If the same beneficial owner appears via multiple chains, percentages are summed
 * and all chains are recorded.
 *
 * Used for: DOR org charts (beneficial ownership disclosure), FinCEN BOI reporting.
 */
export function computeBeneficialOwnership(
  subjectType: SubjectType,
  subjectId: string,
  allOwnership: Ownership[],
  allOwners: Owner[]
): BeneficialOwner[] {
  const accumulator = new Map<string, BeneficialOwner>();

  function walk(
    subjType: SubjectType,
    subjId: string,
    cumulativePct: number,
    pathSoFar: { owner: Owner; relationship: Ownership }[],
    visited: Set<string>
  ) {
    const direct = getDirectOwnersOf(subjType, subjId, allOwnership, allOwners);

    for (const { relationship, owner } of direct) {
      if (!owner) continue;
      const ownerId = String(owner.id);
      if (visited.has(ownerId)) continue;

      const directPct = relationship.fields.OwnershipPercent ?? 0;
      const effectivePct = (cumulativePct * directPct) / 100;

      const ownerType = owner.fields.OwnerType;
      const isTerminal = ownerType === 'Individual' || ownerType === 'Nonprofit';

      if (isTerminal) {
        const existing = accumulator.get(ownerId);
        const path: BeneficialPath = {
          intermediates: [...pathSoFar],
          terminalRelationship: relationship,
          pathPercent: effectivePct,
        };
        if (existing) {
          existing.beneficialPercent += effectivePct;
          existing.paths.push(path);
        } else {
          accumulator.set(ownerId, {
            owner,
            beneficialPercent: effectivePct,
            paths: [path],
          });
        }
      } else {
        // LLC or unknown — recurse upstream
        const newVisited = new Set(visited);
        newVisited.add(ownerId);
        walk('owner', ownerId, effectivePct, [...pathSoFar, { owner, relationship }], newVisited);
      }
    }
  }

  walk(subjectType, subjectId, 100, [], new Set());

  return Array.from(accumulator.values()).sort(
    (a, b) => b.beneficialPercent - a.beneficialPercent
  );
}
