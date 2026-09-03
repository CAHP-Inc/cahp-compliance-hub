import { useState, useMemo } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import { useSharePointList, LIST_NAMES, isCahpEntity, type Property, type PropertyStatus, type CahpState, type Submittal, type SubmittalStatusValue, type SubmittalReview, type TaxMapID, type Contact, type Owner, type Ownership, type OutstandingItem } from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { formatDateET } from '../lib/dates';
import { REVIEW_INTERVAL_DAYS } from '../components/SubmittalReviewsSection';

// A submittal stops needing weekly reviews once it reaches a closed state —
// mirrors CLOSED_STATUSES in SubmittalReviewsSection.tsx / CLOSED_REVIEW in
// PropertyDetail.tsx's SubmittalsTab.
const CLOSED_REVIEW_STATUSES = new Set<SubmittalStatusValue>(['Approved', 'Invoiced', 'Paid', 'Denied', 'Withdrawn']);

// A submittal isn't due for another weekly review until 6 days have passed
// since it was last reviewed (never-reviewed is always due).
const REVIEW_DUE_MS = 6 * 24 * 60 * 60 * 1000;

/** One entity in the nested Properties tree. Children sit below this entity in the
 *  corporate hierarchy; directProperties are properties whose primary direct owner
 *  IS this entity. Both can be empty (an entity might exist only as a passthrough
 *  parent between a top-level holder and the entity that actually owns a property). */
interface EntityNode {
  ownerId: string;
  owner: Owner | null;
  directProperties: Property[];
  children: EntityNode[];
}

const STATUS_STYLES: Record<PropertyStatus, string> = {
  Active: 'bg-green-100 text-green-800',
  Pending: 'bg-yellow-100 text-yellow-800',
  Withdrawn: 'bg-gray-100 text-gray-700',
  'Removed from Program': 'bg-red-100 text-red-800',
  Sold: 'bg-blue-100 text-blue-800',
};

const FILING_STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  'Draft': 'bg-gray-100 text-gray-800',
  'Package Mailed (NC)': 'bg-indigo-100 text-indigo-800',
  'Filed': 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-amber-100 text-amber-800',
  'Responded - Awaiting DOR': 'bg-purple-100 text-purple-800',
  'Approved': 'bg-green-100 text-green-800',
  'Invoiced': 'bg-teal-100 text-teal-800',
  'Paid': 'bg-emerald-100 text-emerald-900',
  'Denied': 'bg-red-100 text-red-800',
  'Withdrawn': 'bg-gray-100 text-gray-500',
};

export function Properties() {
  const navigate = useNavigate();
  const [searchParams, setSearchParams] = useSearchParams();
  const { data, loading, error, refetch } = useSharePointList<Property>(LIST_NAMES.Properties, {
    top: 200,
  });
  const submittals = useSharePointList<Submittal>(LIST_NAMES.Submittals, { top: 500 });
  const submittalReviews = useSharePointList<SubmittalReview>(LIST_NAMES.SubmittalReviews, { top: 2000 });
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const contacts = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 1000 });
  const outstanding = useSharePointList<OutstandingItem>(LIST_NAMES.Outstanding, { top: 2000 });
  const [expandedOwnerIds, setExpandedOwnerIds] = useState<Set<string>>(new Set());

  // Tags the property-detail navigation so its "Back to Properties" link can
  // return here with the Pending Weekly Review view (and its sort) restored,
  // instead of dropping back to the unfiltered default list.
  const goToProperty = (propertyId: string) =>
    navigate(pendingReviewOnly ? `/properties/${propertyId}?from=pending-review` : `/properties/${propertyId}`);

  const toggleExpand = (ownerId: string) =>
    setExpandedOwnerIds((prev) => {
      const next = new Set(prev);
      if (next.has(ownerId)) next.delete(ownerId);
      else next.add(ownerId);
      return next;
    });

  // For each property, find its primary direct owner (largest-percent member).
  // Properties with no direct-owner Ownership row fall into an "ungrouped" bucket.
  const primaryOwnerByProperty = useMemo(() => {
    const map = new Map<string, Owner>();
    const ownersById = new Map<string, Owner>();
    (owners.data ?? []).forEach((o) => ownersById.set(String(o.id), o));
    const rowsByProperty = new Map<string, Ownership[]>();
    (ownership.data ?? []).forEach((row) => {
      const pid = row.fields.LinkedPropertyLookupId ? String(row.fields.LinkedPropertyLookupId) : '';
      if (!pid || !row.fields.OwnerLookupId) return;
      if (!rowsByProperty.has(pid)) rowsByProperty.set(pid, []);
      rowsByProperty.get(pid)!.push(row);
    });
    for (const [pid, rows] of rowsByProperty) {
      const sorted = [...rows].sort(
        (a, b) => (b.fields.OwnershipPercent ?? 0) - (a.fields.OwnershipPercent ?? 0),
      );
      const oid = String(sorted[0].fields.OwnerLookupId);
      const owner = ownersById.get(oid);
      if (owner) map.set(pid, owner);
    }
    return map;
  }, [ownership.data, owners.data]);

  // For each owner, find their parent in the corporate hierarchy. Used to
  // nest the Properties listing arbitrarily deep (Stan -> IV Fund Global ->
  // IV 3 LLC -> Property all roll up into one tree under Stan).
  //
  // Schema: a row "Holdings owns N% of Fund I" has OwnerLookupId=Holdings,
  // ParentOwnerLookupId=Fund I. So to find Fund I's parents, filter by
  // ParentOwnerLookupId=Fund I and look at OwnerLookupId.
  //
  // What counts as a parent: the largest-% holder of this entity, INCLUDING
  // individuals (so Stan can sit at the top of his own chain). CAHP-flagged
  // entities are explicitly skipped — CAHP SC LLC at 0.01% is on every
  // property, but it isn't the operating parent the user wants to group by.
  const parentLLCByOwner = useMemo(() => {
    const map = new Map<string, Owner>();
    const ownersById = new Map<string, Owner>();
    (owners.data ?? []).forEach((o) => ownersById.set(String(o.id), o));
    const rowsByHeldEntity = new Map<string, Ownership[]>();
    (ownership.data ?? []).forEach((row) => {
      const heldId = row.fields.ParentOwnerLookupId ? String(row.fields.ParentOwnerLookupId) : '';
      if (!heldId || !row.fields.OwnerLookupId) return;
      if (!rowsByHeldEntity.has(heldId)) rowsByHeldEntity.set(heldId, []);
      rowsByHeldEntity.get(heldId)!.push(row);
    });
    for (const [heldId, rows] of rowsByHeldEntity) {
      const sorted = [...rows].sort(
        (a, b) => (b.fields.OwnershipPercent ?? 0) - (a.fields.OwnershipPercent ?? 0),
      );
      for (const row of sorted) {
        const oid = String(row.fields.OwnerLookupId);
        const owner = ownersById.get(oid);
        if (!owner) continue;
        // Skip CAHP-flagged entities entirely — they're a compliance affiliate,
        // not an operating parent. Without this, CAHP nonprofit ends up as the
        // root of every property in the system because CAHP SC LLC sits on
        // every property at 0.01%.
        if (isCahpEntity(owner)) continue;
        map.set(heldId, owner);
        break;
      }
    }
    return map;
  }, [ownership.data, owners.data]);

  // contactId → contact, for quick lookup when rendering the Owner Contact column
  const contactsById = useMemo(() => {
    const m = new Map<string, Contact>();
    (contacts.data ?? []).forEach((c) => m.set(String(c.id), c));
    return m;
  }, [contacts.data]);

  /**
   * Parcels per property — used to indicate multi-parcel filings.
   */
  const parcelCountByProperty = useMemo(() => {
    const map = new Map<string, number>();
    (taxMapIDs.data ?? []).forEach((t) => {
      const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
      if (!pid) return;
      map.set(pid, (map.get(pid) ?? 0) + 1);
    });
    return map;
  }, [taxMapIDs.data]);

  /**
   * Per-property searchable parcel text — every linked tax map ID's number and
   * physical address, concatenated and lowercased. Lets the property search box
   * match on a parcel street address (or parcel number) and surface its property.
   */
  const parcelSearchByProperty = useMemo(() => {
    const map = new Map<string, string>();
    (taxMapIDs.data ?? []).forEach((t) => {
      const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
      if (!pid) return;
      const text = `${t.fields.Title ?? ''} ${t.fields.ParcelAddress ?? ''}`.toLowerCase();
      map.set(pid, `${map.get(pid) ?? ''} ${text}`);
    });
    return map;
  }, [taxMapIDs.data]);

  /**
   * Per-property parcel + filed-parcel counts. A parcel is "filed" if at
   * least one submittal points at it (via TaxMapIDLookupId) with a status
   * other than Draft / blank. Lets the entity row surface IV-Fund-style SFR
   * portfolios at a glance: "120 parcels · 87 filed".
   */
  const parcelStatsByProperty = useMemo(() => {
    const filedTaxMapIds = new Set<string>();
    (submittals.data ?? []).forEach((s) => {
      const tmid = s.fields.TaxMapIDLookupId ? String(s.fields.TaxMapIDLookupId) : '';
      if (!tmid) return;
      const status = s.fields.SubmittalStatus;
      if (status && status !== 'Draft') filedTaxMapIds.add(tmid);
    });
    const map = new Map<string, { totalParcels: number; filedParcels: number; sahaParcels: number }>();
    (taxMapIDs.data ?? []).forEach((t) => {
      const pid = t.fields.LinkedPropertyLookupId ? String(t.fields.LinkedPropertyLookupId) : '';
      if (!pid) return;
      const cur = map.get(pid) ?? { totalParcels: 0, filedParcels: 0, sahaParcels: 0 };
      cur.totalParcels++;
      if (filedTaxMapIds.has(String(t.id))) cur.filedParcels++;
      if (t.fields.PriorSAHAAbatement) cur.sahaParcels++;
      map.set(pid, cur);
    });
    return map;
  }, [taxMapIDs.data, submittals.data]);

  /**
   * Per-property open Outstanding Items count. "Open" excludes closed statuses
   * (Done / Received / Not Applicable). Also tracks overdue (DueDate < now and
   * still open) so the column can flag properties that need attention.
   */
  const openItemsByProperty = useMemo(() => {
    const closed = new Set(['Done', 'Received', 'Not Applicable']);
    const now = Date.now();
    const map = new Map<string, { open: number; overdue: number }>();
    (outstanding.data ?? []).forEach((i) => {
      const pid = i.fields.PropertyLookupId ? String(i.fields.PropertyLookupId) : '';
      if (!pid) return;
      if (closed.has(i.fields.ItemStatus ?? '')) return;
      const cur = map.get(pid) ?? { open: 0, overdue: 0 };
      cur.open++;
      if (i.fields.DueDate && new Date(i.fields.DueDate).getTime() < now) {
        cur.overdue++;
      }
      map.set(pid, cur);
    });
    return map;
  }, [outstanding.data]);

  /**
   * For each property's latest year+filing-type, count submittals by status.
   * Used to show "X of N Filed" style multi-parcel breakdown.
   */
  const filingAggregateByProperty = useMemo(() => {
    if (!submittals.data) return new Map<string, { total: number; filed: number; approved: number; denied: number; draft: number; year?: string; filingType?: string }>();
    // Find max-year + tiebreaking latest filing per property first
    const latestKey = new Map<string, { year: string; filingType: string }>();
    submittals.data.forEach((s) => {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      if (!pid) return;
      const year = s.fields.cahpTaxYear ?? '';
      const ft = s.fields.FilingType ?? '';
      const cur = latestKey.get(pid);
      if (!cur || (year > cur.year) || (year === cur.year && ft > cur.filingType)) {
        latestKey.set(pid, { year, filingType: ft });
      }
    });
    // Aggregate counts for that latest year+filing-type
    const result = new Map<string, { total: number; filed: number; approved: number; denied: number; draft: number; year?: string; filingType?: string }>();
    latestKey.forEach((key, pid) => {
      const matching = submittals.data!.filter(
        (s) =>
          String(s.fields.PropertyLookupId ?? '') === pid &&
          (s.fields.cahpTaxYear ?? '') === key.year &&
          (s.fields.FilingType ?? '') === key.filingType
      );
      const agg = {
        total: matching.length,
        filed: 0,
        approved: 0,
        denied: 0,
        draft: 0,
        year: key.year || undefined,
        filingType: key.filingType || undefined,
      };
      matching.forEach((s) => {
        const status = s.fields.SubmittalStatus;
        if (status === 'Draft') agg.draft++;
        // Invoiced / Paid are past Approved — still count as approved here.
        else if (status === 'Approved' || status === 'Invoiced' || status === 'Paid') agg.approved++;
        else if (status === 'Denied') agg.denied++;
        else if (status) agg.filed++; // Filed, Letter Received, Responded etc.
      });
      result.set(pid, agg);
    });
    return result;
  }, [submittals.data]);

  /**
   * Most recent weekly-review timestamp per submittal (ms since epoch), from
   * the Submittal Reviews list. Same "latest by createdDateTime" logic as
   * SubmittalReviewsSection / PropertyDetail's SubmittalsTab.
   */
  const lastReviewBySubmittalId = useMemo(() => {
    const map = new Map<string, number>();
    (submittalReviews.data ?? []).forEach((r) => {
      const sid = r.fields.ReviewSubmittalLookupId ? String(r.fields.ReviewSubmittalLookupId) : '';
      if (!sid) return;
      const t = new Date(r.createdDateTime).getTime();
      if (Number.isNaN(t)) return;
      const prev = map.get(sid);
      if (prev === undefined || t > prev) map.set(sid, t);
    });
    return map;
  }, [submittalReviews.data]);

  /**
   * Properties with at least one submittal that's still "in flight" — filed
   * but not yet Draft (hasn't started) or closed out (Approved/Invoiced/
   * Paid/Denied/Withdrawn) — AND not reviewed within the last 6 days (or
   * never reviewed). These are the ones weekly review actually needs to
   * touch right now. oldestReviewTs is the earliest last-reviewed timestamp
   * across that property's due submittals (never-reviewed counts as
   * -Infinity, i.e. most overdue) — used to sort the neediest properties first.
   */
  const pendingReviewByProperty = useMemo(() => {
    const map = new Map<string, { count: number; oldestReviewTs: number }>();
    const now = Date.now();
    (submittals.data ?? []).forEach((s) => {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      const status = s.fields.SubmittalStatus;
      if (!pid || !status || status === 'Draft' || CLOSED_REVIEW_STATUSES.has(status)) return;
      const lastReviewTs = lastReviewBySubmittalId.get(String(s.id)) ?? -Infinity;
      if (now - lastReviewTs < REVIEW_DUE_MS) return;
      const cur = map.get(pid);
      if (!cur) {
        map.set(pid, { count: 1, oldestReviewTs: lastReviewTs });
      } else {
        cur.count++;
        cur.oldestReviewTs = Math.min(cur.oldestReviewTs, lastReviewTs);
      }
    });
    return map;
  }, [submittals.data, lastReviewBySubmittalId]);

  const [search, setSearch] = useState('');
  const [stateFilter, setStateFilter] = useState<CahpState | 'All'>('All');
  const [statusFilter, setStatusFilter] = useState<PropertyStatus | 'All'>('All');
  // Filing status = the property's most-recent submittal status ('None' = no submittal yet).
  const [filingFilter, setFilingFilter] = useState<SubmittalStatusValue | 'All' | 'None'>('All');
  // Owner contact filter — value is the Contact's listItem ID (string) or 'All' / 'None'.
  const [contactFilter, setContactFilter] = useState<string>('All');
  const [sahaFilter, setSahaFilter] = useState<'All' | 'SAHA' | 'NonSAHA'>('All');
  // Deep-linkable "Pending Weekly Review" view (?view=pending-review) — narrows
  // to properties with an in-flight submittal and defaults the sort to oldest
  // review first, so the queue always surfaces what needs attention next.
  const [pendingReviewOnly, setPendingReviewOnly] = useState(() => searchParams.get('view') === 'pending-review');
  const [sortBy, setSortBy] = useState<'name' | 'filingStatus' | 'reviewDate'>(
    () => (searchParams.get('view') === 'pending-review' ? 'reviewDate' : 'name'),
  );

  const togglePendingReviewOnly = () => {
    setPendingReviewOnly((prev) => {
      const next = !prev;
      if (next) {
        setSortBy('reviewDate');
        setSearchParams({ view: 'pending-review' }, { replace: true });
      } else {
        setSearchParams({}, { replace: true });
      }
      return next;
    });
  };

  /**
   * Build a map of propertyId → most recent submittal for the property.
   * "Most recent" = highest tax year, then latest DateFiled as a tiebreaker.
   * This lets us surface each property's current filing posture on the table.
   */
  const latestSubmittalByProperty = useMemo(() => {
    if (!submittals.data) return new Map<string, Submittal>();
    const map = new Map<string, Submittal>();
    submittals.data.forEach((s) => {
      const pid = s.fields.PropertyLookupId ? String(s.fields.PropertyLookupId) : '';
      if (!pid) return;
      const existing = map.get(pid);
      if (!existing) {
        map.set(pid, s);
        return;
      }
      const existingYear = Number(existing.fields.cahpTaxYear ?? 0);
      const newYear = Number(s.fields.cahpTaxYear ?? 0);
      if (newYear > existingYear) {
        map.set(pid, s);
      } else if (newYear === existingYear) {
        const existingDate = existing.fields.DateFiled ?? '';
        const newDate = s.fields.DateFiled ?? '';
        if (newDate > existingDate) {
          map.set(pid, s);
        }
      }
    });
    return map;
  }, [submittals.data]);

  const filtered = useMemo(() => {
    if (!data) return [];
    const result = data.filter((p) => {
      const f = p.fields;
      if (search) {
        const q = search.toLowerCase();
        const hay = `${f.Title ?? ''} ${f.LegalEntity ?? ''} ${f.PropertyAddress ?? ''}`.toLowerCase();
        const parcelHay = parcelSearchByProperty.get(String(p.id)) ?? '';
        if (!hay.includes(q) && !parcelHay.includes(q)) return false;
      }
      if (stateFilter !== 'All' && f.cahpState !== stateFilter) return false;
      if (statusFilter !== 'All' && f.PropertyStatus !== statusFilter) return false;
      if (filingFilter !== 'All') {
        const fs = latestSubmittalByProperty.get(p.id)?.fields.SubmittalStatus;
        if (filingFilter === 'None' ? !!fs : fs !== filingFilter) return false;
      }
      if (sahaFilter !== 'All') {
        const hasSaha = (parcelStatsByProperty.get(String(p.id))?.sahaParcels ?? 0) > 0;
        if (sahaFilter === 'SAHA' && !hasSaha) return false;
        if (sahaFilter === 'NonSAHA' && hasSaha) return false;
      }
      if (contactFilter !== 'All') {
        const propContactId = f.PropertyOwnerContactLookupId ? String(f.PropertyOwnerContactLookupId) : '';
        if (contactFilter === 'None') {
          if (propContactId) return false;
        } else if (propContactId !== contactFilter) {
          return false;
        }
      }
      if (pendingReviewOnly && !pendingReviewByProperty.has(String(p.id))) return false;
      return true;
    });
    // Client-side sort — SharePoint won't sort server-side because Title isn't indexed
    if (sortBy === 'reviewDate') {
      // Oldest (or never) reviewed first — properties with nothing pending sort last.
      return result.sort((a, b) => {
        const aTs = pendingReviewByProperty.get(String(a.id))?.oldestReviewTs ?? Infinity;
        const bTs = pendingReviewByProperty.get(String(b.id))?.oldestReviewTs ?? Infinity;
        if (aTs !== bTs) return aTs - bTs;
        return (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '');
      });
    }
    if (sortBy === 'filingStatus') {
      // Status order per workflow: Draft → Filed → Letter Received → Responded → Denied → Approved → Withdrawn
      // Properties without any submittal go to the bottom.
      const ORDER: Record<string, number> = {
        'Draft': 1,
        'Filed': 2,
        'Letter Received - Action Needed': 3,
        'Letter Received': 3,
        'Responded - Awaiting DOR': 4,
        'Denied': 5,
        'Approved': 6,
        'Invoiced': 7,
        'Paid': 8,
        'Withdrawn': 9,
      };
      return result.sort((a, b) => {
        const aSub = latestSubmittalByProperty.get(a.id);
        const bSub = latestSubmittalByProperty.get(b.id);
        const aStatus = aSub?.fields.SubmittalStatus ?? '';
        const bStatus = bSub?.fields.SubmittalStatus ?? '';
        const aRank = ORDER[aStatus] ?? 99;
        const bRank = ORDER[bStatus] ?? 99;
        if (aRank !== bRank) return aRank - bRank;
        // Tiebreaker: property name
        return (a.fields.Title ?? '').localeCompare(b.fields.Title ?? '');
      });
    }
    return result.sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''));
  }, [data, search, stateFilter, statusFilter, filingFilter, contactFilter, sahaFilter, pendingReviewOnly, sortBy, latestSubmittalByProperty, parcelSearchByProperty, parcelStatsByProperty, pendingReviewByProperty]);

  const stats = useMemo(() => {
    if (!data) return null;
    let totalParcels = 0;
    let filedParcels = 0;
    let sahaParcels = 0;
    let openItems = 0;
    let overdueItems = 0;
    for (const p of data) {
      const ps = parcelStatsByProperty.get(p.id);
      if (ps) {
        totalParcels += ps.totalParcels;
        filedParcels += ps.filedParcels;
        sahaParcels += ps.sahaParcels;
      }
      const oi = openItemsByProperty.get(p.id);
      if (oi) {
        openItems += oi.open;
        overdueItems += oi.overdue;
      }
    }
    return {
      total: data.length,
      active: data.filter((p) => p.fields.PropertyStatus === 'Active').length,
      sc: data.filter((p) => p.fields.cahpState === 'SC').length,
      nc: data.filter((p) => p.fields.cahpState === 'NC').length,
      units: data.reduce((sum, p) => sum + (p.fields.UnitCount ?? 0), 0),
      parcels: totalParcels,
      filedParcels,
      sahaParcels,
      openItems,
      overdueItems,
    };
  }, [data, parcelStatsByProperty, openItemsByProperty]);

  /**
   * Two-level nested grouping:
   *   - Outer group = topmost-LLC ancestor of each property's primary direct owner
   *     (or the primary owner itself if it has no LLC parent in the system).
   *   - Inner sub-group = the primary direct owner, when that owner has a parent LLC
   *     and is therefore nested.
   *
   * So VanRock Holdings ends up as a top-level group containing both its 7
   * direct-ownership properties AND a sub-group "VanRock Fund I, LLC" with
   * its own properties (701 E Main).
   *
   * Properties with no direct owner fall into an "Unlinked" bucket at the
   * bottom.
   */
  /**
   * Recursive N-level nested grouping. Walks the parent chain upward from
   * every property's primary direct owner until it hits an entity with no
   * known parent — that entity becomes a top-level group. Sub-entities
   * nest underneath their parent, with arbitrary depth (so Stan → IV Fund
   * Global → IV 3 LLC → Property all collapses into one expandable tree).
   *
   * Unlinked properties (no direct-owner Ownership row) drop into an
   * UNLINKED bucket at the bottom of the list.
   */
  const groupedRows = useMemo(() => {
    const UNLINKED = '__unlinked__';

    // 1. Determine which entities are in scope: any entity that primary-owns
    //    a filtered property OR sits anywhere above one in the parent chain.
    const inScope = new Set<string>();
    const directPropsByOwner = new Map<string, Property[]>();
    const unlinkedProps: Property[] = [];
    for (const p of filtered) {
      const primary = primaryOwnerByProperty.get(String(p.id));
      if (!primary) {
        unlinkedProps.push(p);
        continue;
      }
      const arr = directPropsByOwner.get(String(primary.id)) ?? [];
      arr.push(p);
      directPropsByOwner.set(String(primary.id), arr);
      // Walk upward, capping depth to avoid runaway loops on malformed data.
      let cur: Owner | undefined = primary;
      const guard = new Set<string>();
      let depth = 0;
      while (cur && !guard.has(String(cur.id)) && depth < 12) {
        guard.add(String(cur.id));
        inScope.add(String(cur.id));
        cur = parentLLCByOwner.get(String(cur.id));
        depth++;
      }
    }

    // 2. Build a node per in-scope entity, parented to its parentLLC (if that
    //    parent is also in scope) or marked as a root.
    const ownersById = new Map<string, Owner>();
    (owners.data ?? []).forEach((o) => ownersById.set(String(o.id), o));

    const nodeById = new Map<string, EntityNode>();
    for (const id of inScope) {
      const owner = ownersById.get(id) ?? null;
      nodeById.set(id, {
        ownerId: id,
        owner,
        directProperties: directPropsByOwner.get(id) ?? [],
        children: [],
      });
    }
    const roots: EntityNode[] = [];
    for (const node of nodeById.values()) {
      const parent = parentLLCByOwner.get(node.ownerId);
      const parentNode = parent ? nodeById.get(String(parent.id)) : undefined;
      if (parentNode) parentNode.children.push(node);
      else roots.push(node);
    }

    // 3. Sort alphabetically by Title at every level. Unlinked group is
    //    appended as a synthetic root at the very end.
    const sortRec = (nodes: EntityNode[]) => {
      nodes.sort((a, b) =>
        (a.owner?.fields.Title ?? '').localeCompare(b.owner?.fields.Title ?? ''),
      );
      for (const n of nodes) sortRec(n.children);
    };
    sortRec(roots);

    if (unlinkedProps.length > 0) {
      roots.push({
        ownerId: UNLINKED,
        owner: null,
        directProperties: unlinkedProps,
        children: [],
      });
    }
    return roots;
  }, [filtered, primaryOwnerByProperty, parentLLCByOwner, owners.data]);

  // When a search or filter narrows the list, force every owner group open so
  // matches aren't hidden inside a collapsed group (a parcel-address match often
  // lives under a multi-property entity that's collapsed by default).
  const anyFilterActive =
    !!search.trim() || stateFilter !== 'All' || statusFilter !== 'All' || contactFilter !== 'All' || sahaFilter !== 'All' || pendingReviewOnly;

  if (loading) return <LoadingState />;
  if (error) return <ErrorState error={error} onRetry={refetch} />;
  if (!data || !stats) return null;

  return (
    <div>
      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold text-teal-700">Properties</h1>
          <p className="text-sm text-gray-500 mt-1">
            {stats.total} properties under CAHP management · {stats.units.toLocaleString()} units · {stats.parcels.toLocaleString()} tax map IDs ({stats.filedParcels} filed{stats.sahaParcels > 0 ? `, ${stats.sahaParcels} SAHA` : ''})
          </p>
        </div>
        <button
          onClick={() => navigate('/properties/new')}
          className="bg-teal-700 hover:bg-teal-900 text-white px-4 py-2 rounded-md font-medium flex items-center gap-2 transition-colors"
        >
          <Icon name="plus" size={16} />
          New Property
        </button>
      </div>

      {/* KPI cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-4 mb-6">
        <KPICard label="Total Properties" value={stats.total} />
        <KPICard label="Active" value={stats.active} accent="success" />
        <KPICard label="Total Units" value={stats.units.toLocaleString()} />
        <KPICard label="Tax Map IDs" value={stats.parcels.toLocaleString()} />
        <KPICard
          label="Filed Parcels"
          value={stats.parcels > 0 ? `${stats.filedParcels} / ${stats.parcels}` : '—'}
          accent={stats.parcels > 0 && stats.filedParcels === stats.parcels ? 'success' : undefined}
        />
        <KPICard
          label={stats.overdueItems > 0 ? `Open Items (${stats.overdueItems} overdue)` : 'Open Items'}
          value={stats.openItems.toLocaleString()}
          accent={stats.overdueItems > 0 ? 'error' : undefined}
        />
        <KPICard label="SC" value={stats.sc} />
        <KPICard label="NC" value={stats.nc} />
      </div>

      {/* Filters */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 flex flex-wrap gap-2 items-center shadow-card">
        <button
          onClick={togglePendingReviewOnly}
          title="Show only properties with a submittal due for weekly review (in flight, and not reviewed in the last 6 days), oldest review first"
          className={`text-xs px-3 py-1.5 rounded-md font-semibold border transition-colors flex items-center gap-1.5 whitespace-nowrap ${
            pendingReviewOnly
              ? 'bg-amber-600 border-amber-600 text-white'
              : 'bg-white border-gray-200 text-gray-700 hover:border-amber-300'
          }`}
        >
          <Icon name="history" size={14} />
          Pending Weekly Review
          <span className={`px-1.5 rounded-full text-[10px] font-mono-data ${pendingReviewOnly ? 'bg-white/20' : 'bg-amber-100 text-amber-800'}`}>
            {pendingReviewByProperty.size}
          </span>
        </button>
        <div className="relative flex-1 min-w-[240px]">
          <Icon name="search" size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search property, entity, address, parcel # / parcel address…"
            className="w-full pl-9 pr-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500"
          />
        </div>
        <Select value={stateFilter} onChange={(v) => setStateFilter(v as CahpState | 'All')} options={['All', 'SC', 'NC']} />
        <Select
          value={statusFilter}
          onChange={(v) => setStatusFilter(v as PropertyStatus | 'All')}
          options={['All', 'Active', 'Pending', 'Withdrawn', 'Removed from Program', 'Sold']}
        />
        <select
          value={filingFilter}
          onChange={(e) => setFilingFilter(e.target.value as SubmittalStatusValue | 'All' | 'None')}
          className="text-xs px-2 py-1.5 border border-gray-200 rounded-md bg-white focus:outline-none focus:border-teal-500"
          title="Filter by filing status (most recent submittal)"
        >
          <option value="All">All filing statuses</option>
          <option value="None">— No submittal yet —</option>
          <option value="Draft">Draft</option>
          <option value="Filed">Filed</option>
          <option value="Letter Received - Action Needed">Letter Received - Action Needed</option>
          <option value="Responded - Awaiting DOR">Responded - Awaiting DOR</option>
          <option value="Approved">Approved</option>
          <option value="Invoiced">Invoiced</option>
          <option value="Paid">Paid</option>
          <option value="Denied">Denied</option>
          <option value="Withdrawn">Withdrawn</option>
          <option value="Package Mailed (NC)">Package Mailed (NC)</option>
        </select>
        <select
          value={contactFilter}
          onChange={(e) => setContactFilter(e.target.value)}
          className="text-xs px-2 py-1.5 border border-gray-200 rounded-md bg-white focus:outline-none focus:border-teal-500"
          title="Filter by owner contact"
        >
          <option value="All">All contacts</option>
          <option value="None">— No contact set —</option>
          {[...(contacts.data ?? [])]
            .sort((a, b) => (a.fields.Title ?? '').localeCompare(b.fields.Title ?? ''))
            .map((c) => (
              <option key={c.id} value={String(c.id)}>{c.fields.Title}</option>
            ))}
        </select>
        <select
          value={sahaFilter}
          onChange={(e) => setSahaFilter(e.target.value as 'All' | 'SAHA' | 'NonSAHA')}
          className="text-xs px-2 py-1.5 border border-gray-200 rounded-md bg-white focus:outline-none focus:border-teal-500"
          title="Filter by prior SAHA abatement"
        >
          <option value="All">All parcels (SAHA)</option>
          <option value="SAHA">SAHA only</option>
          <option value="NonSAHA">Exclude SAHA</option>
        </select>
        <div className="flex items-center gap-1">
          <span className="text-[11px] text-gray-500 uppercase tracking-wider font-semibold">Sort:</span>
          <select
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value as 'name' | 'filingStatus' | 'reviewDate')}
            className="text-xs px-2 py-1.5 border border-gray-200 rounded-md bg-white focus:outline-none focus:border-teal-500"
          >
            <option value="name">Property name</option>
            <option value="filingStatus">Filing status</option>
            <option value="reviewDate">Oldest review</option>
          </select>
        </div>
        {filtered.length !== data.length && (
          <span className="text-xs text-gray-500 px-1">
            {filtered.length} of {data.length}
          </span>
        )}
      </div>

      {/* Table */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left w-8"></th>
                <th className="px-4 py-3 text-left">Legal Entity</th>
                <th className="px-4 py-3 text-left">Property</th>
                <th className="px-4 py-3 text-left">State</th>
                <th className="px-4 py-3 text-left">County</th>
                <th className="px-4 py-3 text-right">Units</th>
                <th className="px-4 py-3 text-left">AMI</th>
                <th className="px-4 py-3 text-left">Owner Contact</th>
                <th className="px-4 py-3 text-right" title="Open Outstanding Items">Open</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Filing Status</th>
                {pendingReviewOnly && (
                  <th className="px-4 py-3 text-left" title="Oldest last-reviewed date among this property's in-flight submittals">Last Reviewed</th>
                )}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {(() => {
                // ─── Cell renderers shared across all row variants ───
                const renderPropertyNameCell = (p: Property) => (
                  <>
                    {p.fields.Title}
                    {p.fields.cahpVerificationStatus === 'Inherited - Unverified' && (
                      <span className="ml-2 inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold bg-yellow-50 text-yellow-800 align-middle">UNVERIFIED</span>
                    )}
                    {p.fields.cahpVerificationStatus === 'Needs Follow-Up' && (
                      <span className="ml-2 inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold bg-error/10 text-error align-middle">FOLLOW-UP</span>
                    )}
                  </>
                );
                const renderCountyCell = (p: Property) => {
                  const list = (p.fields.cahpCounty ?? '').split(',').map((s) => s.trim()).filter(Boolean);
                  if (list.length === 0) return '—';
                  const stripped = list.map((c) => c.replace(/\s*\([^)]*\)\s*/g, ''));
                  if (stripped.length === 1) return stripped[0];
                  return (
                    <div className="flex flex-wrap gap-1">
                      {stripped.map((c) => (
                        <span key={c} className="px-1 py-0.5 rounded bg-teal-50 text-teal-800 text-[10px] font-medium">{c}</span>
                      ))}
                    </div>
                  );
                };
                const renderContactCell = (p: Property) => {
                  const cId = p.fields.PropertyOwnerContactLookupId ? String(p.fields.PropertyOwnerContactLookupId) : '';
                  const contact = cId ? contactsById.get(cId) : undefined;
                  if (!contact) return <span className="text-gray-400">—</span>;
                  return (
                    <div className="min-w-0">
                      <div className="text-gray-900 truncate">{contact.fields.Title}</div>
                      {contact.fields.ContactEmail && (
                        <div className="text-[11px] text-gray-500 font-mono-data truncate">{contact.fields.ContactEmail}</div>
                      )}
                    </div>
                  );
                };
                const renderOpenItemsCell = (open: number, overdue: number) => {
                  if (open === 0) return <span className="text-gray-300">—</span>;
                  return (
                    <div className="flex flex-col items-end" title={`${open} open${overdue > 0 ? ` (${overdue} overdue)` : ''}`}>
                      <span className={`font-mono-data font-semibold ${overdue > 0 ? 'text-error' : 'text-gray-800'}`}>{open}</span>
                      {overdue > 0 && (
                        <span className="text-[10px] text-error font-mono-data">{overdue} overdue</span>
                      )}
                    </div>
                  );
                };
                const renderUnitsCell = (p: Property) => {
                  const ps = parcelStatsByProperty.get(p.id);
                  return (
                    <div className="flex flex-col items-end">
                      <span className="font-mono-data">{p.fields.UnitCount ?? '—'}</span>
                      {ps && ps.totalParcels > 0 && (
                        <span className="text-[10px] text-gray-500 font-mono-data whitespace-nowrap" title={`${ps.filedParcels} of ${ps.totalParcels} parcels filed${ps.sahaParcels > 0 ? ` · ${ps.sahaParcels} prior SAHA abatement` : ''}`}>
                          {ps.totalParcels} TMID · <span className={ps.filedParcels === ps.totalParcels ? 'text-success' : 'text-gray-600'}>{ps.filedParcels} filed</span>
                          {ps.sahaParcels > 0 && <> · <span className="text-gold-700">{ps.sahaParcels} SAHA</span></>}
                        </span>
                      )}
                    </div>
                  );
                };
                const renderFilingStatusCell = (p: Property) => {
                  const agg = filingAggregateByProperty.get(p.id);
                  const sub = latestSubmittalByProperty.get(p.id);
                  const parcelCount = parcelCountByProperty.get(p.id) ?? 0;
                  if (!sub || !agg) {
                    return (
                      <div className="flex flex-col gap-0.5 items-start">
                        <span className="text-gray-400 text-xs italic">Not Filed</span>
                        {parcelCount > 1 && (
                          <span className="text-[10px] text-gray-400 font-mono-data">{parcelCount} parcels</span>
                        )}
                      </div>
                    );
                  }
                  const status = sub.fields.SubmittalStatus;
                  const year = agg.year ?? sub.fields.cahpTaxYear;
                  const filingType = agg.filingType ?? sub.fields.FilingType;
                  const multiParcel = agg.total > 1;
                  const allApproved = agg.approved === agg.total;
                  const allDenied = agg.denied === agg.total;
                  const allDraft = agg.draft === agg.total;
                  const allFiled = agg.filed === agg.total;
                  const isMixed = multiParcel && !allApproved && !allDenied && !allDraft && !allFiled;
                  const headlineStatus = isMixed ? 'Mixed' : status;
                  return (
                    <div className="flex flex-col gap-0.5 items-start">
                      {headlineStatus && (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${FILING_STATUS_STYLES[headlineStatus as SubmittalStatusValue] || 'bg-purple-100 text-purple-800'}`}>{headlineStatus}</span>
                      )}
                      {(year || filingType) && (
                        <span className="text-[10px] text-gray-500 font-mono-data">{year ?? ''}{year && filingType ? ' · ' : ''}{filingType ?? ''}</span>
                      )}
                      {multiParcel && (
                        <span className="text-[10px] text-gray-600 font-mono-data" title={`${agg.draft} Draft / ${agg.filed} Filed / ${agg.approved} Approved / ${agg.denied} Denied`}>
                          {isMixed ? (
                            <>
                              {agg.approved > 0 && <span className="text-green-700">{agg.approved}A</span>}
                              {agg.approved > 0 && (agg.filed + agg.denied + agg.draft > 0) && ' / '}
                              {agg.filed > 0 && <span className="text-blue-700">{agg.filed}F</span>}
                              {agg.filed > 0 && (agg.denied + agg.draft > 0) && ' / '}
                              {agg.denied > 0 && <span className="text-red-700">{agg.denied}D</span>}
                              {agg.denied > 0 && agg.draft > 0 && ' / '}
                              {agg.draft > 0 && <span className="text-gray-600">{agg.draft}Dr</span>}
                              {' of '}{agg.total}
                            </>
                          ) : (
                            <>{agg.total} of {parcelCount} parcels</>
                          )}
                        </span>
                      )}
                    </div>
                  );
                };
                // Renders a "Last Reviewed" cell from an oldest-review timestamp
                // (ms since epoch, -Infinity = never reviewed, null = nothing pending).
                const renderLastReviewedCell = (ts: number | null) => {
                  if (ts === null) return <span className="text-gray-300">—</span>;
                  if (ts === -Infinity) {
                    return <span className="inline-block px-1.5 py-0.5 rounded text-[11px] font-semibold bg-red-100 text-red-800">Never reviewed</span>;
                  }
                  const days = Math.floor((Date.now() - ts) / 86400000);
                  return (
                    <div className="flex flex-col items-start">
                      <span className="font-mono-data">{formatDateET(new Date(ts))}</span>
                      <span className={`text-[10px] font-mono-data ${days >= REVIEW_INTERVAL_DAYS ? 'text-error' : 'text-gray-500'}`}>
                        {days === 0 ? 'today' : `${days}d ago`}
                      </span>
                    </div>
                  );
                };
                // Oldest last-reviewed timestamp across a set of properties, considering
                // only the ones with an in-flight submittal. Null if none are pending.
                const groupOldestReviewTs = (props: Property[]): number | null => {
                  const tsList = props
                    .map((p) => pendingReviewByProperty.get(String(p.id))?.oldestReviewTs)
                    .filter((v): v is number => v !== undefined);
                  if (tsList.length === 0) return null;
                  return Math.min(...tsList);
                };

                // Collect every property below a node (direct + recursively from children).
                const collectAllProps = (n: EntityNode): Property[] => {
                  const out = [...n.directProperties];
                  for (const c of n.children) out.push(...collectAllProps(c));
                  return out;
                };

                // Property rows nest under their owning entity. Padding scales with
                // depth so a property under Stan → IV Fund Global → IV 3 LLC sits
                // visually further in than one under just Holdings → 701 E Main.
                const renderPropertyRow = (p: Property, depth: number, viaSubEntity?: string) => {
                  const oi = openItemsByProperty.get(p.id) ?? { open: 0, overdue: 0 };
                  // 32px base + 24px per nesting level (depth 1 = under top group)
                  const padLeft = 32 + Math.max(0, depth - 1) * 24;
                  return (
                    <tr
                      key={`prop-${depth}-${p.id}`}
                      onClick={() => goToProperty(p.id)}
                      className="hover:bg-gray-50 transition-colors cursor-pointer"
                    >
                      <td className="px-4 py-3"></td>
                      <td className="px-4 py-3 text-gray-400 text-xs" style={{ paddingLeft: padLeft }}>
                        ↳ {viaSubEntity && <span className="text-gray-500 italic ml-1">via {viaSubEntity}</span>}
                      </td>
                      <td className="px-4 py-3 font-medium text-gray-900">{renderPropertyNameCell(p)}</td>
                      <td className="px-4 py-3">
                        <span className="font-mono-data text-xs font-semibold text-teal-700">{p.fields.cahpState || '—'}</span>
                      </td>
                      <td className="px-4 py-3 text-gray-700 text-xs">{renderCountyCell(p)}</td>
                      <td className="px-4 py-3 text-right">{renderUnitsCell(p)}</td>
                      <td className="px-4 py-3 text-gray-700 text-xs">{p.fields.AMIProgram || '—'}</td>
                      <td className="px-4 py-3 text-xs">{renderContactCell(p)}</td>
                      <td className="px-4 py-3 text-right">{renderOpenItemsCell(oi.open, oi.overdue)}</td>
                      <td className="px-4 py-3">
                        {p.fields.PropertyStatus ? (
                          <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[p.fields.PropertyStatus] || 'bg-gray-100 text-gray-700'}`}>{p.fields.PropertyStatus}</span>
                        ) : '—'}
                      </td>
                      <td className="px-4 py-3">{renderFilingStatusCell(p)}</td>
                      {pendingReviewOnly && (
                        <td className="px-4 py-3 text-xs">{renderLastReviewedCell(pendingReviewByProperty.get(String(p.id))?.oldestReviewTs ?? null)}</td>
                      )}
                    </tr>
                  );
                };

                // Render an entity (top-level OR sub-entity, depth-styled).
                // Returns the rows for this entity + recursively its children
                // when expanded. A leaf entity with exactly one property and no
                // children collapses to a single flat row.
                const renderEntity = (node: EntityNode, depth: number): JSX.Element[] => {
                  const ownerKey = node.ownerId;
                  const ownerName = node.owner?.fields.Title ?? '(no linked owner)';
                  const isExpanded = anyFilterActive || expandedOwnerIds.has(ownerKey);
                  const allProps = collectAllProps(node);
                  const childCount = node.children.length;

                  // ─── Leaf optimization: single property, no children, at top-level depth ───
                  if (depth === 0 && childCount === 0 && allProps.length === 1) {
                    const p = allProps[0];
                    const oi = openItemsByProperty.get(p.id) ?? { open: 0, overdue: 0 };
                    return [
                      <tr
                        key={`single-${ownerKey}-${p.id}`}
                        onClick={() => goToProperty(p.id)}
                        className="hover:bg-gray-50 transition-colors cursor-pointer"
                      >
                        <td className="px-4 py-3"></td>
                        <td className="px-4 py-3 font-medium text-gray-900">
                          {node.owner ? (
                            <button
                              onClick={(e) => { e.stopPropagation(); navigate(`/owners/${node.owner!.id}`); }}
                              className="text-teal-700 hover:text-teal-900 underline-offset-2 hover:underline text-left"
                            >
                              {ownerName}
                            </button>
                          ) : (
                            <span className="text-gray-400 italic">{ownerName}</span>
                          )}
                        </td>
                        <td className="px-4 py-3 font-medium text-gray-900">{renderPropertyNameCell(p)}</td>
                        <td className="px-4 py-3">
                          <span className="font-mono-data text-xs font-semibold text-teal-700">{p.fields.cahpState || '—'}</span>
                        </td>
                        <td className="px-4 py-3 text-gray-700 text-xs">{renderCountyCell(p)}</td>
                        <td className="px-4 py-3 text-right">{renderUnitsCell(p)}</td>
                        <td className="px-4 py-3 text-gray-700 text-xs">{p.fields.AMIProgram || '—'}</td>
                        <td className="px-4 py-3 text-xs">{renderContactCell(p)}</td>
                        <td className="px-4 py-3 text-right">{renderOpenItemsCell(oi.open, oi.overdue)}</td>
                        <td className="px-4 py-3">
                          {p.fields.PropertyStatus ? (
                            <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${STATUS_STYLES[p.fields.PropertyStatus] || 'bg-gray-100 text-gray-700'}`}>{p.fields.PropertyStatus}</span>
                          ) : '—'}
                        </td>
                        <td className="px-4 py-3">{renderFilingStatusCell(p)}</td>
                        {pendingReviewOnly && (
                          <td className="px-4 py-3 text-xs">{renderLastReviewedCell(pendingReviewByProperty.get(String(p.id))?.oldestReviewTs ?? null)}</td>
                        )}
                      </tr>,
                    ];
                  }

                  // ─── Multi-property / has children: aggregate row + expanded contents ───
                  const totalUnits = allProps.reduce((s, p) => s + (p.fields.UnitCount ?? 0), 0);
                  const totalParcels = allProps.reduce((s, p) => s + (parcelStatsByProperty.get(p.id)?.totalParcels ?? 0), 0);
                  const filedParcels = allProps.reduce((s, p) => s + (parcelStatsByProperty.get(p.id)?.filedParcels ?? 0), 0);
                  const sahaParcels = allProps.reduce((s, p) => s + (parcelStatsByProperty.get(p.id)?.sahaParcels ?? 0), 0);
                  const totalOpenItems = allProps.reduce((s, p) => s + (openItemsByProperty.get(p.id)?.open ?? 0), 0);
                  const totalOverdueItems = allProps.reduce((s, p) => s + (openItemsByProperty.get(p.id)?.overdue ?? 0), 0);
                  const stateAgg = countBy(allProps, (p) => p.fields.cahpState);
                  const countyAgg = countBy(
                    allProps.flatMap((p) => (p.fields.cahpCounty ?? '').split(',').map((s) => s.trim()).filter(Boolean).map((c) => ({ c }))),
                    (x) => x.c,
                  );
                  const amiAgg = countBy(allProps, (p) => p.fields.AMIProgram);
                  const statusAgg = countBy(allProps, (p) => p.fields.PropertyStatus);
                  const filingAgg = countBy(
                    allProps.map((p) => {
                      const sub = latestSubmittalByProperty.get(p.id);
                      return { s: sub?.fields.SubmittalStatus ?? 'Not Filed' };
                    }),
                    (x) => x.s,
                  );
                  const contactAgg = countBy(
                    allProps.map((p) => {
                      const cId = p.fields.PropertyOwnerContactLookupId ? String(p.fields.PropertyOwnerContactLookupId) : '';
                      return { c: contactsById.get(cId)?.fields.Title };
                    }),
                    (x) => x.c,
                  );
                  const groupOldestTs = groupOldestReviewTs(allProps);

                  // Styling scales with depth: depth 0 is the bold top-level row;
                  // depth >0 is a lighter sub-entity row with progressive indent.
                  const rowBg = depth === 0 ? 'bg-gray-50/50' : 'bg-amber-50/30';
                  const padLeft = 16 + depth * 24;
                  const titleClass = depth === 0 ? 'font-bold text-gray-900' : 'text-gray-700 font-semibold text-xs';

                  const rows: JSX.Element[] = [];
                  rows.push(
                    <tr
                      key={`group-${ownerKey}-${depth}`}
                      onClick={() => toggleExpand(ownerKey)}
                      className={`hover:bg-gray-50 transition-colors cursor-pointer ${rowBg}`}
                    >
                      <td className="px-4 py-3 text-gray-500">
                        {depth === 0 && (
                          <Icon name="chevron-right" size={14} className={isExpanded ? 'rotate-90 transition-transform' : 'transition-transform'} />
                        )}
                      </td>
                      <td className={`px-4 py-3 ${titleClass}`} style={{ paddingLeft: padLeft }}>
                        <div className="flex items-center gap-1">
                          {depth > 0 && (
                            <Icon name="chevron-right" size={12} className={isExpanded ? 'rotate-90 transition-transform' : 'transition-transform'} />
                          )}
                          {node.owner ? (
                            <button
                              onClick={(e) => { e.stopPropagation(); navigate(`/owners/${node.owner!.id}`); }}
                              className={`text-teal-700 hover:text-teal-900 underline-offset-2 hover:underline text-left ${depth === 0 ? '' : 'text-xs'}`}
                            >
                              {ownerName}
                            </button>
                          ) : (
                            <span className="text-gray-400 italic">{ownerName}</span>
                          )}
                          {childCount > 0 && (
                            <span className="ml-2 text-[10px] text-gray-500 font-normal">
                              + {childCount} sub-{childCount === 1 ? 'entity' : 'entities'}
                            </span>
                          )}
                        </div>
                      </td>
                      <td className="px-4 py-3 text-xs text-gray-600">
                        {allProps.length} {allProps.length === 1 ? 'property' : 'properties'}
                      </td>
                      <td className="px-4 py-3"><AggregateChips entries={stateAgg} /></td>
                      <td className="px-4 py-3">{depth === 0 ? <AggregateChips entries={countyAgg} styleMap={{}} /> : null}</td>
                      <td className="px-4 py-3 text-right">
                        <div className="flex flex-col items-end">
                          <span className={`font-mono-data ${depth === 0 ? 'font-semibold' : 'text-xs font-semibold'}`}>{totalUnits || '—'}</span>
                          {totalParcels > 0 && (
                            <span className="text-[10px] text-gray-500 font-mono-data whitespace-nowrap" title={`${filedParcels} of ${totalParcels} parcels filed${sahaParcels > 0 ? ` · ${sahaParcels} prior SAHA abatement` : ''}`}>
                              {totalParcels} TMID · <span className={filedParcels === totalParcels ? 'text-success' : 'text-gray-600'}>{filedParcels} filed</span>
                              {sahaParcels > 0 && <> · <span className="text-gold-700">{sahaParcels} SAHA</span></>}
                            </span>
                          )}
                        </div>
                      </td>
                      <td className="px-4 py-3">{depth === 0 ? <AggregateChips entries={amiAgg} /> : null}</td>
                      <td className="px-4 py-3">{depth === 0 ? <AggregateChips entries={contactAgg} /> : null}</td>
                      <td className="px-4 py-3 text-right">{renderOpenItemsCell(totalOpenItems, totalOverdueItems)}</td>
                      <td className="px-4 py-3"><AggregateChips entries={statusAgg} styleMap={STATUS_STYLES as Record<string, string>} /></td>
                      <td className="px-4 py-3"><AggregateChips entries={filingAgg} styleMap={FILING_STATUS_STYLES as Record<string, string>} /></td>
                      {pendingReviewOnly && (
                        <td className="px-4 py-3 text-xs">{renderLastReviewedCell(groupOldestTs)}</td>
                      )}
                    </tr>,
                  );

                  if (isExpanded) {
                    // Direct properties owned by this entity first
                    for (const p of node.directProperties) {
                      rows.push(renderPropertyRow(p, depth + 1, depth > 0 ? ownerName : undefined));
                    }
                    // Then recurse into sub-entities
                    for (const child of node.children) {
                      rows.push(...renderEntity(child, depth + 1));
                    }
                  }
                  return rows;
                };

                return groupedRows.flatMap((node) => renderEntity(node, 0));
              })()}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={pendingReviewOnly ? 12 : 11} className="px-4 py-8 text-center text-gray-500 text-sm">
                    No properties match the current filters.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <p className="text-xs text-gray-400 mt-4 text-center">
        Click any property to drill into its detail page · Click <strong>New Property</strong> to add one.
      </p>
    </div>
  );
}

/**
 * Render distinct values from a set of properties as small badge chips with
 * counts. Used in the collapsed aggregate row.
 *   countBy([{x:'a'},{x:'a'},{x:'b'}], p => p.x)  →  [{label:'a',count:2},{label:'b',count:1}]
 */
function countBy<T>(items: T[], key: (item: T) => string | undefined): { label: string; count: number }[] {
  const m = new Map<string, number>();
  for (const it of items) {
    const k = (key(it) ?? '').trim();
    if (!k) continue;
    m.set(k, (m.get(k) ?? 0) + 1);
  }
  return Array.from(m.entries()).map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count);
}

function AggregateChips({
  entries,
  emptyLabel = '—',
  styleMap,
}: {
  entries: { label: string; count: number }[];
  emptyLabel?: string;
  /** Optional per-label background class, e.g. STATUS_STYLES. */
  styleMap?: Record<string, string>;
}) {
  if (entries.length === 0) return <span className="text-gray-300">{emptyLabel}</span>;
  return (
    <div className="flex flex-wrap gap-1">
      {entries.map((e) => (
        <span
          key={e.label}
          className={`px-1.5 py-0.5 rounded text-[10px] font-semibold inline-flex items-center gap-1 ${styleMap?.[e.label] ?? 'bg-gray-100 text-gray-700'}`}
        >
          <span className="truncate max-w-[120px]">{e.label}</span>
          <span className="font-mono-data opacity-75">{e.count}</span>
        </span>
      ))}
    </div>
  );
}

function KPICard({ label, value, accent }: { label: string; value: string | number; accent?: 'success' | 'error' }) {
  const accentClass =
    accent === 'success' ? 'text-success' :
    accent === 'error' ? 'text-error' :
    'text-teal-700';
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
      <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">{label}</div>
      <div className={`text-3xl font-bold mt-1 ${accentClass}`}>{value}</div>
    </div>
  );
}

function Select({ value, onChange, options }: { value: string; onChange: (v: string) => void; options: readonly string[] }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="px-3 py-1.5 border border-gray-200 rounded-md text-sm focus:outline-none focus:border-teal-500 bg-white"
    >
      {options.map((o) => (
        <option key={o} value={o}>
          {o}
        </option>
      ))}
    </select>
  );
}

function LoadingState() {
  return (
    <div>
      <h1 className="text-2xl font-bold text-teal-700 mb-6">Properties</h1>
      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center gap-3 text-gray-500">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin"></div>
          <span className="text-sm">Loading properties from SharePoint…</span>
        </div>
      </div>
    </div>
  );
}

function ErrorState({ error, onRetry }: { error: Error; onRetry: () => void }) {
  return (
    <div>
      <h1 className="text-2xl font-bold text-teal-700 mb-6">Properties</h1>
      <div className="bg-red-50 border border-red-200 rounded-lg p-6">
        <div className="font-semibold text-error mb-2 flex items-center gap-2">
          <Icon name="alert" size={18} />
          Failed to load properties
        </div>
        <p className="text-sm text-red-700 mb-3 font-mono-data">{error.message}</p>
        <div className="text-xs text-red-600 space-y-1">
          <p>Common causes:</p>
          <ul className="list-disc list-inside ml-2 space-y-0.5">
            <li>
              <code className="font-mono-data">VITE_SHAREPOINT_SITE</code> env var is missing or
              wrong (should be <code>vanrockre.sharepoint.com:/sites/CAHPComplianceHub</code>)
            </li>
            <li>
              The list name doesn't match exactly — expecting <code>Properties Registry</code>
            </li>
            <li>Insufficient SharePoint permissions on the site</li>
            <li>MSAL token couldn't be acquired silently for SharePoint scopes</li>
          </ul>
        </div>
        <button
          onClick={onRetry}
          className="mt-4 text-sm text-teal-700 hover:text-teal-900 font-medium underline"
        >
          Retry
        </button>
      </div>
    </div>
  );
}
