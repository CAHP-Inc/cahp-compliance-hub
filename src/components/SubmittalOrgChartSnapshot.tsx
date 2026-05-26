import { useMemo, useState } from 'react';
import { Icon } from './ui/Icon';
import { formatDateTime } from '../lib/dates';

/**
 * Renders an org chart from a frozen JSON snapshot stored on a submittal.
 *
 * Snapshot shape (per `serializeTree` in SubmittalDetail.tsx):
 * [
 *   {
 *     ownerId, ownerTitle, ownerType,
 *     relationshipType, ownershipPercent, effectiveDate,
 *     children: [...]
 *   }
 * ]
 *
 * This view is read-only and historical — it does NOT reflect current ownership.
 * Spec §3.6.6: "The submittal's Org Chart sub-view always renders from the snapshot, not from current data."
 */

// =============================================================================
// Snapshot types
// =============================================================================

export interface SnapshotNode {
  ownerId?: string;
  ownerTitle?: string;
  ownerType?: string;
  relationshipType?: string;
  ownershipPercent?: number;
  effectiveDate?: string;
  children: SnapshotNode[];
}

export interface SnapshotEnvelope {
  version: number;
  capturedAt: string;
  propertyId?: string;
  tree: SnapshotNode[];
}

interface BeneficialOwnerFromSnapshot {
  ownerId?: string;
  ownerTitle?: string;
  ownerType?: string;
  beneficialPercent: number;
  paths: { intermediates: string[]; pathPercent: number }[];
}

// =============================================================================
// Component
// =============================================================================

type ChartLayout = 'detailed' | 'beneficial' | 'dor';

const LAYOUT_INFO: Record<ChartLayout, { label: string; description: string }> = {
  detailed: {
    label: 'Detailed',
    description: 'Full chain as captured at filing — property at top, members below.',
  },
  beneficial: {
    label: 'Beneficial',
    description: 'Compounded percentages to natural-person and nonprofit owners as of the filing date.',
  },
  dor: {
    label: 'DOR-Friendly',
    description: 'Property at the bottom per DOR convention. Same data, inverted.',
  },
};

const OWNER_TYPE_BADGE_STYLES: Record<string, string> = {
  Individual: 'bg-blue-100 text-blue-800',
  LLC: 'bg-purple-100 text-purple-800',
  Nonprofit: 'bg-teal-100 text-teal-800',
  Trust: 'bg-amber-100 text-amber-800',
  Corporation: 'bg-indigo-100 text-indigo-800',
  'Limited Partnership': 'bg-rose-100 text-rose-800',
  'General Partnership': 'bg-fuchsia-100 text-fuchsia-800',
};

export function SubmittalOrgChartSnapshot({
  snapshotJSON,
  capturedAt,
  propertyTitle,
}: {
  snapshotJSON: string | undefined;
  capturedAt: string | undefined;
  propertyTitle: string;
}) {
  const [layout, setLayout] = useState<ChartLayout>('detailed');

  const snapshot = useMemo<SnapshotEnvelope | null>(() => {
    if (!snapshotJSON) return null;
    try {
      const parsed = JSON.parse(snapshotJSON);
      // Tolerate both raw tree array and full envelope
      if (Array.isArray(parsed)) {
        return { version: 0, capturedAt: capturedAt ?? '', tree: parsed };
      }
      return parsed as SnapshotEnvelope;
    } catch {
      return null;
    }
  }, [snapshotJSON, capturedAt]);

  const beneficial = useMemo(() => {
    if (!snapshot) return [];
    return computeBeneficialFromSnapshot(snapshot.tree);
  }, [snapshot]);

  // ─── Empty state: no snapshot yet ───
  if (!snapshotJSON) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100">
          <h3 className="text-sm font-semibold text-gray-700">Org Chart Snapshot</h3>
          <p className="text-xs text-gray-500 mt-0.5">
            Snapshot is captured automatically on the <strong>Draft → Filed</strong> transition.
          </p>
        </div>
        <div className="px-4 py-8 text-center">
          <Icon name="folder" size={32} className="text-gray-300 mx-auto mb-2" />
          <p className="text-sm text-gray-500">No snapshot yet — this submittal hasn't been filed.</p>
          <p className="text-xs text-gray-400 mt-1">
            When you transition to Filed, the entire ownership chain is frozen as JSON on this submittal.
            Future ownership edits will not change what's shown here.
          </p>
        </div>
      </div>
    );
  }

  // ─── Snapshot JSON failed to parse ───
  if (!snapshot) {
    return (
      <div className="bg-red-50 border border-red-200 rounded-lg p-4">
        <div className="font-semibold text-error mb-1 text-sm">Snapshot data corrupted</div>
        <p className="text-xs text-red-700">
          The snapshot JSON on this submittal didn't parse cleanly. Raw value preserved in SharePoint
          (Submittals Tracker → OrgChartSnapshotJSON column).
        </p>
      </div>
    );
  }

  const isEmpty = snapshot.tree.length === 0;

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100 flex items-start justify-between flex-wrap gap-2">
        <div>
          <h3 className="text-sm font-semibold text-gray-700 flex items-center gap-2">
            <Icon name="check" size={14} className="text-success" />
            Org Chart Snapshot
            <span className="text-[10px] font-semibold bg-gold-100 text-gold-900 px-1.5 py-0.5 rounded uppercase tracking-wider">
              Frozen
            </span>
          </h3>
          <p className="text-xs text-gray-500 mt-0.5">
            Captured{' '}
            <span className="font-mono-data">
              {snapshot.capturedAt
                ? formatDateTime(snapshot.capturedAt)
                : 'date unknown'}
            </span>
            . Renders historical data only — not current ownership.
          </p>
        </div>
      </div>

      {/* Layout switcher */}
      <div className="px-4 py-2 border-b border-gray-100 bg-gray-50">
        <div className="flex flex-wrap gap-2 mb-1">
          {(Object.keys(LAYOUT_INFO) as ChartLayout[]).map((key) => (
            <button
              key={key}
              onClick={() => setLayout(key)}
              className={`px-2 py-1 rounded text-xs font-medium transition-colors ${
                layout === key ? 'bg-teal-700 text-white' : 'bg-white text-gray-700 hover:bg-gray-100 border border-gray-200'
              }`}
            >
              {LAYOUT_INFO[key].label}
            </button>
          ))}
        </div>
        <p className="text-[11px] text-gray-500">{LAYOUT_INFO[layout].description}</p>
      </div>

      {/* Content */}
      <div className="p-5">
        {isEmpty ? (
          <p className="text-sm text-gray-500 italic text-center py-6">
            Snapshot captured but ownership tree was empty. No direct owners of this property at the filing date.
          </p>
        ) : (
          <>
            {layout === 'detailed' && <DetailedView tree={snapshot.tree} propertyTitle={propertyTitle} />}
            {layout === 'beneficial' && <BeneficialView beneficial={beneficial} propertyTitle={propertyTitle} />}
            {layout === 'dor' && <DORView tree={snapshot.tree} propertyTitle={propertyTitle} />}
          </>
        )}
      </div>
    </div>
  );
}

// =============================================================================
// Layout: Detailed (top-down chain)
// =============================================================================

function DetailedView({ tree, propertyTitle }: { tree: SnapshotNode[]; propertyTitle: string }) {
  return (
    <div>
      <PropertyRootNode title={propertyTitle} />
      <div className="pl-6 ml-3 border-l-2 border-gray-300 mt-2 space-y-2">
        {tree.map((node, idx) => (
          <SnapshotBranch key={`${node.ownerId ?? idx}-${idx}`} node={node} />
        ))}
      </div>
    </div>
  );
}

function SnapshotBranch({ node }: { node: SnapshotNode }) {
  return (
    <div>
      <EntityCard
        name={node.ownerTitle ?? '(unresolved)'}
        ownerType={node.ownerType}
        relationshipType={node.relationshipType}
        percent={node.ownershipPercent}
      />
      {node.children.length > 0 && (
        <div className="pl-6 ml-3 border-l-2 border-gray-300 mt-2 space-y-2">
          {node.children.map((child, idx) => (
            <SnapshotBranch key={`${child.ownerId ?? idx}-${idx}`} node={child} />
          ))}
        </div>
      )}
    </div>
  );
}

// =============================================================================
// Layout: Beneficial (compounded to terminals)
// =============================================================================

function BeneficialView({
  beneficial,
  propertyTitle,
}: {
  beneficial: BeneficialOwnerFromSnapshot[];
  propertyTitle: string;
}) {
  const totalTraced = beneficial.reduce((sum, b) => sum + b.beneficialPercent, 0);

  if (beneficial.length === 0) {
    return (
      <p className="text-sm text-gray-500 italic text-center py-6">
        No natural-person or nonprofit terminal owners in the snapshot. Chain may not have been fully populated at filing time.
      </p>
    );
  }

  return (
    <div>
      <div className="mb-3 text-xs text-gray-600">
        <strong>{beneficial.length}</strong> beneficial owner{beneficial.length === 1 ? '' : 's'} of {propertyTitle}
        {' · '}
        <span className="font-mono-data">{totalTraced.toFixed(2)}% traced</span>
        {totalTraced < 99.9 && (
          <span className="text-warning ml-1">
            · {(100 - totalTraced).toFixed(2)}% not traced (chain incomplete at snapshot time)
          </span>
        )}
      </div>
      <ul className="divide-y divide-gray-100">
        {beneficial.map((b, idx) => (
          <li key={`${b.ownerId ?? idx}-${idx}`} className="py-2">
            <div className="flex items-center justify-between gap-3">
              <div className="flex items-center gap-2 flex-1 min-w-0">
                <span className="font-medium text-gray-900 truncate">{b.ownerTitle}</span>
                {b.ownerType && (
                  <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold flex-shrink-0 ${OWNER_TYPE_BADGE_STYLES[b.ownerType] ?? 'bg-gray-100 text-gray-700'}`}>
                    {b.ownerType}
                  </span>
                )}
              </div>
              <span className="font-mono-data text-sm font-semibold text-teal-700 flex-shrink-0">
                {b.beneficialPercent.toFixed(2)}%
              </span>
            </div>
            <div className="mt-1.5 h-1.5 bg-gray-100 rounded-full overflow-hidden">
              <div
                className="h-full bg-teal-500 rounded-full"
                style={{ width: `${Math.min(100, b.beneficialPercent)}%` }}
              />
            </div>
            {b.paths.length > 0 && (
              <div className="mt-1.5 text-xs text-gray-500 font-mono-data">
                {b.paths.map((path, pidx) => (
                  <div key={pidx}>
                    via {path.intermediates.length === 0 ? '(direct)' : path.intermediates.join(' ← ')}
                    {' · '}{path.pathPercent.toFixed(2)}%
                  </div>
                ))}
              </div>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
}

// =============================================================================
// Layout: DOR-Friendly (property at bottom)
// =============================================================================

function DORView({ tree, propertyTitle }: { tree: SnapshotNode[]; propertyTitle: string }) {
  const levelGroups = useMemo(() => groupByDepth(tree), [tree]);
  const maxDepth = levelGroups.length;

  return (
    <div className="overflow-x-auto">
      <p className="text-xs text-gray-500 mb-4 italic">
        Beneficial owners at the top, property at the bottom — the orientation DOR prefers for submissions.
      </p>
      <div className="flex flex-col items-center space-y-3 min-w-fit">
        {[...levelGroups].reverse().map((nodes, idx) => (
          <div key={idx} className="flex flex-col items-center">
            <div className="flex flex-wrap justify-center gap-3 max-w-4xl">
              {nodes.map((node, nidx) => (
                <div key={`${node.ownerId ?? nidx}-${nidx}`} className="min-w-[200px]">
                  <EntityCard
                    name={node.ownerTitle ?? '(unresolved)'}
                    ownerType={node.ownerType}
                    relationshipType={node.relationshipType}
                    percent={node.ownershipPercent}
                  />
                </div>
              ))}
            </div>
            {idx < maxDepth && (
              <div className="my-1 text-gray-400 text-xs flex flex-col items-center">
                <div className="w-0.5 h-3 bg-gray-300" />
              </div>
            )}
          </div>
        ))}
        <div className="min-w-[200px] pt-1">
          <PropertyRootNode title={propertyTitle} />
        </div>
      </div>
    </div>
  );
}

function groupByDepth(tree: SnapshotNode[]): SnapshotNode[][] {
  const groups: SnapshotNode[][] = [];
  function walk(nodes: SnapshotNode[], depth: number) {
    if (!groups[depth]) groups[depth] = [];
    nodes.forEach((n) => {
      groups[depth].push(n);
      if (n.children.length > 0) walk(n.children, depth + 1);
    });
  }
  walk(tree, 0);
  return groups;
}

// =============================================================================
// Shared node renderers
// =============================================================================

function PropertyRootNode({ title }: { title: string }) {
  return (
    <div className="inline-block bg-gold-50 border-2 border-gold-500 rounded-lg px-4 py-2 shadow-sm">
      <div className="flex items-center gap-2">
        <Icon name="folder" size={14} className="text-gold-700" />
        <span className="font-bold text-teal-900 text-sm">{title}</span>
        <span className="inline-block px-2 py-0.5 rounded text-[10px] font-semibold bg-gold-200 text-gold-900">
          PROPERTY
        </span>
      </div>
    </div>
  );
}

function EntityCard({
  name,
  ownerType,
  relationshipType,
  percent,
}: {
  name: string;
  ownerType?: string;
  relationshipType?: string;
  percent?: number;
}) {
  return (
    <div className="inline-block bg-white border border-gray-300 rounded-lg px-3 py-2 shadow-sm">
      <div className="flex items-center gap-2 flex-wrap">
        <span className="font-semibold text-gray-900 text-sm">{name}</span>
        {ownerType && (
          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold ${OWNER_TYPE_BADGE_STYLES[ownerType] ?? 'bg-gray-100 text-gray-700'}`}>
            {ownerType}
          </span>
        )}
      </div>
      <div className="text-xs text-gray-600 mt-0.5 font-mono-data">
        {relationshipType ?? 'Member'} · {percent != null ? `${percent}%` : '—'}
      </div>
    </div>
  );
}

// =============================================================================
// Beneficial owner compute from snapshot
// =============================================================================

function computeBeneficialFromSnapshot(tree: SnapshotNode[]): BeneficialOwnerFromSnapshot[] {
  const accumulator = new Map<string, BeneficialOwnerFromSnapshot>();

  function walk(
    nodes: SnapshotNode[],
    cumulativePct: number,
    pathSoFar: string[]
  ) {
    for (const node of nodes) {
      const directPct = node.ownershipPercent ?? 0;
      const effectivePct = (cumulativePct * directPct) / 100;
      const isTerminal = node.ownerType === 'Individual' || node.ownerType === 'Nonprofit';

      if (isTerminal) {
        const key = node.ownerId ?? node.ownerTitle ?? Math.random().toString();
        const existing = accumulator.get(key);
        const pathEntry = { intermediates: [...pathSoFar], pathPercent: effectivePct };
        if (existing) {
          existing.beneficialPercent += effectivePct;
          existing.paths.push(pathEntry);
        } else {
          accumulator.set(key, {
            ownerId: node.ownerId,
            ownerTitle: node.ownerTitle,
            ownerType: node.ownerType,
            beneficialPercent: effectivePct,
            paths: [pathEntry],
          });
        }
      } else if (node.children.length > 0) {
        walk(node.children, effectivePct, [...pathSoFar, node.ownerTitle ?? '(unnamed)']);
      }
    }
  }

  walk(tree, 100, []);

  return Array.from(accumulator.values()).sort(
    (a, b) => b.beneficialPercent - a.beneficialPercent
  );
}
