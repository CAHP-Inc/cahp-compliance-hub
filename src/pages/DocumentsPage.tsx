import { useMemo } from 'react';
import { Link } from 'react-router-dom';
import {
  useSharePointList,
  type Owner,
} from '../lib/sharepoint';
import { PROPERTY_LINKED_LIBRARIES, CAHP_ENTITY_LIBRARY } from '../components/UploadDocumentModal';
import { Icon } from '../components/ui/Icon';

interface DocItemRaw {
  id: string;
  webUrl?: string;
  fields: {
    Title?: string;
    FileLeafRef?: string;
    PropertyLookupId?: string;
    OwnerLookupId?: string;
    Modified?: string;
  };
  lastModifiedDateTime: string;
}

const LIBRARY_ICONS: Record<string, 'file' | 'folder' | 'star' | 'check' | 'home'> = {
  'AMI Certification Letters': 'check',
  'DOR Correspondence': 'file',
  'DOR Submittal Packages': 'file',
  'Land Use Restriction Agreements': 'home',
  'Operating Agreements': 'folder',
  'Org Charts': 'star',
  'Property Deeds': 'home',
  'Supporting Documentation': 'folder',
  'CAHP Entity Documents': 'star',
};

const LIBRARY_DESCRIPTIONS: Record<string, string> = {
  'AMI Certification Letters': 'Income-restricted housing AMI cert renewals.',
  'DOR Correspondence': 'Letters to/from SC DOR.',
  'DOR Submittal Packages': 'Final filed submittal packages with signatures.',
  'Land Use Restriction Agreements': 'Recorded LURAs binding the property.',
  'Operating Agreements': 'Entity OAs — CAHP, property-owner LLCs, member entities.',
  'Org Charts': 'Visual ownership diagrams per property.',
  'Property Deeds': 'Recorded deeds establishing title.',
  'Supporting Documentation': 'EIN letters, COE, Articles, IRS determinations, rent rolls, anything else.',
  'CAHP Entity Documents': 'Nonprofit + CAHP SC LLC entity-level docs (OAs, formation, EIN, determination letters). Reused across all property filings.',
};

export function DocumentsPage() {
  const owners = useSharePointList<Owner>('Owners', { top: 500 });

  // Fetch each library individually — same pattern as PropertyDocumentsTab
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });
  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];
  // 9th library — dedicated CAHP Entity Documents
  const cahpLib = useSharePointList<DocItemRaw>(CAHP_ENTITY_LIBRARY, { top: 500 });

  const loading = libraries.some((l) => l.loading) || owners.loading || cahpLib.loading;

  // Per-library stats — 8 property-linked libraries + 1 CAHP entity library
  const stats = useMemo(() => {
    const all = libraries.map((lib, idx) => {
      const libraryName: string = PROPERTY_LINKED_LIBRARIES[idx];
      const docs = lib.data ?? [];
      const taggedProperty = docs.filter((d) => d.fields.PropertyLookupId).length;
      const taggedOwner = docs.filter((d) => d.fields.OwnerLookupId && !d.fields.PropertyLookupId).length;
      const untagged = docs.filter((d) => !d.fields.PropertyLookupId && !d.fields.OwnerLookupId).length;
      const thirtyAgo = Date.now() - 30 * 24 * 60 * 60 * 1000;
      const recent = docs.filter((d) => {
        const date = d.fields.Modified ?? d.lastModifiedDateTime;
        return date && new Date(date).getTime() > thirtyAgo;
      }).length;
      return {
        library: libraryName,
        isCahpEntity: false,
        total: docs.length,
        taggedProperty,
        taggedOwner,
        untagged,
        recent,
      };
    });

    // Append the CAHP Entity Documents library — no tagging concept, library IS the tag
    const cahpDocs = cahpLib.data ?? [];
    const thirtyAgo = Date.now() - 30 * 24 * 60 * 60 * 1000;
    const cahpRecent = cahpDocs.filter((d) => {
      const date = d.fields.Modified ?? d.lastModifiedDateTime;
      return date && new Date(date).getTime() > thirtyAgo;
    }).length;
    all.push({
      library: CAHP_ENTITY_LIBRARY,
      isCahpEntity: true,
      total: cahpDocs.length,
      taggedProperty: 0,
      taggedOwner: cahpDocs.length, // every file in here is implicitly tagged to the CAHP entity
      untagged: 0,
      recent: cahpRecent,
    });

    return all;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data, cahpLib.data]);

  const totalDocs = stats.reduce((sum, s) => sum + s.total, 0);
  const totalUntagged = stats.reduce((sum, s) => sum + s.untagged, 0);
  const totalRecent = stats.reduce((sum, s) => sum + s.recent, 0);

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">Documents</h1>
        <p className="text-sm text-gray-500 mt-1">
          Browse the eight SharePoint document libraries that back the CAHP filing workflow.
          Most uploads happen from inside Property → Documents, but this view gives you a portfolio-wide picture.
        </p>
      </div>

      {/* Top-level KPIs */}
      <div className="grid grid-cols-3 gap-4 mb-6">
        <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
          <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Total Documents</div>
          <div className="text-3xl font-bold text-teal-700 mt-1">{loading ? '…' : totalDocs}</div>
          <div className="text-xs text-gray-500 mt-1">Across all 8 libraries</div>
        </div>
        <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
          <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Untagged</div>
          <div className={`text-3xl font-bold mt-1 ${totalUntagged > 0 ? 'text-warning' : 'text-success'}`}>
            {loading ? '…' : totalUntagged}
          </div>
          <div className="text-xs text-gray-500 mt-1">
            <Link to="/untagged-documents" className="text-teal-700 hover:underline">View queue →</Link>
          </div>
        </div>
        <div className="bg-white border border-gray-200 rounded-lg p-4 shadow-card">
          <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Uploaded This Month</div>
          <div className="text-3xl font-bold text-teal-700 mt-1">{loading ? '…' : totalRecent}</div>
          <div className="text-xs text-gray-500 mt-1">Last 30 days</div>
        </div>
      </div>

      {/* Library grid */}
      <h2 className="text-sm font-semibold text-gray-700 uppercase tracking-wider mb-3">Libraries</h2>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        {stats.map((s) => (
          <div
            key={s.library}
            className="bg-white border border-gray-200 rounded-lg p-4 shadow-card hover:shadow-card-hover transition-shadow"
          >
            <div className="flex items-start gap-3 mb-3">
              <Icon name={LIBRARY_ICONS[s.library]} size={18} className="text-teal-700 flex-shrink-0 mt-0.5" />
              <div className="flex-1 min-w-0">
                <h3 className="text-sm font-semibold text-teal-900">{s.library}</h3>
                <p className="text-xs text-gray-500 mt-0.5 leading-snug">{LIBRARY_DESCRIPTIONS[s.library]}</p>
              </div>
              <div className="text-right flex-shrink-0">
                <div className="text-2xl font-bold text-teal-700 font-mono-data">{loading ? '…' : s.total}</div>
                <div className="text-[10px] text-gray-500 uppercase tracking-wider">files</div>
              </div>
            </div>
            <div className="grid grid-cols-4 gap-2 pt-3 border-t border-gray-100">
              {s.isCahpEntity ? (
                <>
                  <div className="col-span-3 text-[10px] text-gray-500 italic self-center">
                    CAHP entity-only library — library membership is the tag, no per-doc tagging.
                  </div>
                  <Stat label="Recent" value={s.recent} />
                </>
              ) : (
                <>
                  <Stat label="Property" value={s.taggedProperty} />
                  <Stat label="Entity" value={s.taggedOwner} />
                  <Stat label="Untagged" value={s.untagged} warn={s.untagged > 0} />
                  <Stat label="Recent" value={s.recent} />
                </>
              )}
            </div>
          </div>
        ))}
      </div>

      <div className="mt-6 bg-blue-50 border border-blue-200 rounded-md p-4 text-xs text-blue-900">
        <strong>How documents flow:</strong> Files are uploaded from <em>Property → Documents</em> tab,
        <em> Owner → Owner Documents</em> section, or <em>CAHP Entity → Documents</em> section. Each upload tags
        the file with either a Property ID or an Owner ID. The <strong>Outstanding Items</strong> Link/Upload
        action lets you fulfill checklist items by attaching the right doc. Untagged files (no Property OR Owner)
        surface in the <Link to="/untagged-documents" className="underline">Untagged Documents</Link> queue for
        cleanup.
      </div>
    </div>
  );
}

function Stat({ label, value, warn }: { label: string; value: number; warn?: boolean }) {
  return (
    <div className="text-center">
      <div className={`text-base font-bold font-mono-data ${warn ? 'text-warning' : 'text-gray-700'}`}>
        {value}
      </div>
      <div className="text-[9px] text-gray-500 uppercase tracking-wider">{label}</div>
    </div>
  );
}
