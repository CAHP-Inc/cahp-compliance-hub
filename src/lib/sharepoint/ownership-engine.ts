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

/**
 * Does this owner count as a "CAHP entity" — i.e., the parent 501(c)(3) or
 * one of its wholly-owned subsidiaries? Used by org chart views to highlight
 * the exemption chain.
 *
 * Honors the explicit `IsCAHPEntity` flag first (Settings → Owner Detail).
 * Falls back to a name-based heuristic so existing data without the flag
 * still gets the right treatment.
 */
export function isCahpEntity(owner: Owner | null | undefined): boolean {
  if (!owner) return false;
  if (owner.fields.IsCAHPEntity) return true;
  const t = (owner.fields.Title ?? '').toLowerCase();
  return t.includes('cahp') || t.includes('carolina affordable housing project');
}

/**
 * Per-node annotations for the property's ownership tree, derived in one
 * pass so the renderer doesn't have to re-walk the chain for each card.
 */
export interface OrgChartCahpAnnotation {
  /** This entity itself is part of the CAHP family. */
  isCahpEntity: boolean;
  /**
   * One of THIS entity's direct upstream members is a CAHP entity — meaning
   * this entity's documents are the ones DOR needs for the exemption
   * filing on its subsidiary properties.
   */
  isExemptionSource: boolean;
  /**
   * Names of the CAHP entities that are direct members of this entity.
   * Empty when isExemptionSource is false.
   */
  cahpMemberNames: string[];
}

/**
 * Walk an OwnershipNode tree and annotate each node with CAHP-related flags.
 * Returned as a Map keyed by relationship.id so the renderer can look up
 * annotations by tree position.
 */
export function annotateCahpChain(
  tree: OwnershipNode[],
): Map<string, OrgChartCahpAnnotation> {
  const annotations = new Map<string, OrgChartCahpAnnotation>();
  function walk(nodes: OwnershipNode[]) {
    for (const node of nodes) {
      const cahpKidsNames: string[] = [];
      for (const parentNode of node.children) {
        if (isCahpEntity(parentNode.owner)) {
          const name = parentNode.owner?.fields.Title;
          if (name && !cahpKidsNames.includes(name)) cahpKidsNames.push(name);
        }
      }
      annotations.set(String(node.relationship.id), {
        isCahpEntity: isCahpEntity(node.owner),
        isExemptionSource: cahpKidsNames.length > 0,
        cahpMemberNames: cahpKidsNames,
      });
      if (node.children.length > 0) walk(node.children);
    }
  }
  walk(tree);
  return annotations;
}

/** Convenience: list every entity in the tree whose direct members include a CAHP entity. */
export function getExemptionSources(
  tree: OwnershipNode[],
): { ownerName: string; cahpMembers: string[] }[] {
  const out: { ownerName: string; cahpMembers: string[] }[] = [];
  const seen = new Set<string>();
  function walk(nodes: OwnershipNode[]) {
    for (const node of nodes) {
      const cahpKidsNames: string[] = [];
      for (const parentNode of node.children) {
        if (isCahpEntity(parentNode.owner)) {
          const name = parentNode.owner?.fields.Title;
          if (name && !cahpKidsNames.includes(name)) cahpKidsNames.push(name);
        }
      }
      if (cahpKidsNames.length > 0 && node.owner?.fields.Title) {
        const key = String(node.owner.id);
        if (!seen.has(key)) {
          seen.add(key);
          out.push({ ownerName: node.owner.fields.Title, cahpMembers: cahpKidsNames });
        }
      }
      if (node.children.length > 0) walk(node.children);
    }
  }
  walk(tree);
  return out;
}

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
        if (ownerType === 'LLC' || ownerType === 'Nonprofit' || ownerType === 'Limited Partnership' || ownerType === 'General Partnership' || ownerType === 'Corporation' || ownerType === 'Trust') {
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
 * Returns the set of ALL owner IDs that sit upstream of a given property —
 * direct owners of the property, parents of those owners, and so on up to
 * the natural-person / nonprofit terminals.
 *
 * Use case: document linking. An outstanding item for a property tagged to
 * CAHP NC LLC may need to pull a document tagged to the parent nonprofit
 * (Carolina Affordable Housing Project Inc) — the chain knows about the
 * relationship; we shouldn't rely on name- or state-based heuristics.
 *
 * Cycle-safe: tracks visited owner IDs so a malformed loop in the chain
 * doesn't recurse forever.
 */
export function getUpstreamOwnerIds(
  propertyId: string,
  allOwnership: Ownership[],
): Set<string> {
  const upstream = new Set<string>();
  const queue: string[] = [];

  // Seed: direct owners of the property
  for (const o of allOwnership) {
    if (
      String(o.fields.LinkedPropertyLookupId) === String(propertyId) &&
      o.fields.OwnerLookupId
    ) {
      const id = String(o.fields.OwnerLookupId);
      if (!upstream.has(id)) {
        upstream.add(id);
        queue.push(id);
      }
    }
  }

  // BFS up the chain: for each owner, find rows where ParentOwner == that owner,
  // and the row's OwnerLookupId is who owns it (one step upstream).
  while (queue.length > 0) {
    const current = queue.shift()!;
    for (const o of allOwnership) {
      if (String(o.fields.ParentOwnerLookupId) !== current) continue;
      if (!o.fields.OwnerLookupId) continue;
      const parent = String(o.fields.OwnerLookupId);
      if (upstream.has(parent)) continue;
      upstream.add(parent);
      queue.push(parent);
    }
  }

  return upstream;
}

/**
 * Returns the set of property IDs this owner has a direct OR indirect beneficial
 * interest in.
 *
 * Direct = rows where OwnerLookupId == ownerId AND LinkedPropertyLookupId is set.
 * Indirect = walk DOWN through child entities (rows where ParentOwnerLookupId
 * points at the current owner) and collect their property holdings, recursively.
 *
 * Use case: "show me everything Deepak owns" — surfaces every property even
 * when Deepak only holds them through nested LLCs.
 */
export function getPropertyIdsForOwner(
  ownerId: string,
  allOwnership: Ownership[],
): Set<string> {
  const propertiesFound = new Set<string>();
  const visited = new Set<string>([String(ownerId)]);

  // Direct holdings
  for (const o of allOwnership) {
    if (
      String(o.fields.OwnerLookupId) === String(ownerId) &&
      o.fields.LinkedPropertyLookupId
    ) {
      propertiesFound.add(String(o.fields.LinkedPropertyLookupId));
    }
  }

  // Indirect: walk down the entity tree
  function walkDown(currentOwnerId: string) {
    const childRelations = allOwnership.filter(
      (o) => String(o.fields.ParentOwnerLookupId) === String(currentOwnerId),
    );
    for (const rel of childRelations) {
      const childOwnerId = rel.fields.OwnerLookupId;
      if (!childOwnerId || visited.has(String(childOwnerId))) continue;
      visited.add(String(childOwnerId));

      for (const o of allOwnership) {
        if (
          String(o.fields.OwnerLookupId) === String(childOwnerId) &&
          o.fields.LinkedPropertyLookupId
        ) {
          propertiesFound.add(String(o.fields.LinkedPropertyLookupId));
        }
      }

      walkDown(String(childOwnerId));
    }
  }
  walkDown(String(ownerId));

  return propertiesFound;
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
