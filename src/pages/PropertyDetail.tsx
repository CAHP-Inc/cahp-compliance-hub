import { useState, useEffect, useMemo } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import {
  useSharePointItem,
  useSharePointList,
  createListItem,
  updateListItem,
  deleteListItem,
  LIST_NAMES,
  type Property,
  type PropertyFields,
  type Submittal,
  type TaxMapID,
  type PropertyStatus,
  type SubmittalStatusValue,
  type ComplianceDeadline,
  type DeadlineStatus,
  type Ownership,
  type Owner,
  type RelationshipType,
  type PropertyNote,
  type Correspondence,
  type Billing,
  type AuditLog,
  type OutstandingItem,
  type ItemStatus,
  type ItemPriority,
  getBeneficialOwnershipTree,
  computeBeneficialOwnership,
  type OwnershipNode,
  type BeneficialOwner,
  type Contact,
} from '../lib/sharepoint';
import { Icon } from '../components/ui/Icon';
import { DispositionModal } from '../components/DispositionModal';
import { DeletePropertyModal } from '../components/DeletePropertyModal';
import { useSession } from '../lib/session';
import { LogLetterModal } from '../components/LogLetterModal';
import { UploadDocumentModal } from '../components/UploadDocumentModal';
import { NewOutstandingItemModal } from '../components/NewOutstandingItemModal';
import { ExportOutstandingItemsModal } from '../components/ExportOutstandingItemsModal';
import { ContactPicker } from '../components/ContactPicker';
import { LinkOrUploadDocumentModal } from '../components/LinkOrUploadDocumentModal';
import { FilingChecklistGenerator } from '../components/FilingChecklistGenerator';
import { EntityDocumentsSection } from '../components/EntityDocumentsSection';
import { EditOwnershipModal } from '../components/EditOwnershipModal';
import { TaxMapIDsSection } from '../components/TaxMapIDsSection';
import { DeedsSection } from '../components/DeedsSection';
import { NewSubmittalModal, BulkCreateSubmittalsModal } from '../components/NewSubmittalModal';
import { formatDateOnly } from '../lib/dates';

const STATUS_STYLES: Record<PropertyStatus, string> = {
  Active: 'bg-green-100 text-green-800 border-green-200',
  Pending: 'bg-yellow-100 text-yellow-800 border-yellow-200',
  Withdrawn: 'bg-gray-100 text-gray-700 border-gray-200',
  'Removed from Program': 'bg-red-100 text-red-800 border-red-200',
  Sold: 'bg-blue-100 text-blue-800 border-blue-200',
};

const SUBMITTAL_STATUS_STYLES: Record<SubmittalStatusValue, string> = {
  Draft: 'bg-gray-100 text-gray-700',
  'Package Mailed (NC)': 'bg-blue-100 text-blue-800',
  Filed: 'bg-blue-100 text-blue-800',
  'Letter Received - Action Needed': 'bg-yellow-100 text-yellow-800',
  'Responded - Awaiting DOR': 'bg-amber-100 text-amber-800',
  Approved: 'bg-green-100 text-green-800',
  Denied: 'bg-red-100 text-red-800',
  Withdrawn: 'bg-gray-100 text-gray-700',
};

// Choice options matching SharePoint config exactly
const CHOICES = {
  AMIProgram: ['20/50', '40/60', '50/80', '60/80', 'Mixed', 'None'] as const,
  CAHPLanguageAdded: ['Yes', 'No', 'In Progress', 'Needs Revision'] as const,
  cahpCounty: [
    'Greenville (SC)', 'Spartanburg (SC)', 'Anderson (SC)', 'Pickens (SC)',
    'Laurens (SC)', 'York (SC)', 'Mecklenburg (NC)', 'Guilford (NC)',
    'Durham (NC)', 'Wake (NC)', 'Forsyth (NC)', 'Buncombe (NC)', 'Other',
  ] as const,
  cahpState: ['SC', 'NC'] as const,
  cahpVerificationStatus: [
    'Inherited - Unverified', 'Verified', 'Needs Follow-Up', 'N/A',
  ] as const,
  LURAExecuted: ['Yes', 'No', 'In Progress', 'N/A'] as const,
  PropertyStatus: ['Active', 'Pending', 'Withdrawn', 'Removed from Program', 'Sold'] as const,
} as const;

type TabId = 'overview' | 'submittals' | 'compliance' | 'ownership' | 'orgChart' | 'correspondence' | 'outstanding' | 'billing' | 'documents' | 'activity' | 'notes';

const TABS: { id: TabId; label: string; shipped: boolean }[] = [
  { id: 'overview', label: 'Overview', shipped: true },
  { id: 'submittals', label: 'Submittals', shipped: true },
  { id: 'compliance', label: 'Compliance', shipped: true },
  { id: 'ownership', label: 'Ownership', shipped: true },
  { id: 'orgChart', label: 'Org Chart', shipped: true },
  { id: 'correspondence', label: 'Correspondence', shipped: true },
  { id: 'outstanding', label: 'Outstanding', shipped: true },
  { id: 'billing', label: 'Billing', shipped: true },
  { id: 'documents', label: 'Documents', shipped: true },
  { id: 'activity', label: 'Activity', shipped: true },
  { id: 'notes', label: 'Notes', shipped: true },
];

export function PropertyDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [activeTab, setActiveTab] = useState<TabId>('overview');

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<PropertyFields | null>(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [dispositionOpen, setDispositionOpen] = useState(false);
  const [deleteOpen, setDeleteOpen] = useState(false);
  const [checklistOpen, setChecklistOpen] = useState(false);
  const [checklistResult, setChecklistResult] = useState<{ created: number; matched: number } | null>(null);

  const session = useSession();
  const isAdmin = session.role === 'Admin';

  const { data: property, loading, error, refetch } = useSharePointItem<Property>(
    LIST_NAMES.Properties, id
  );

  const { data: allSubmittals } = useSharePointList<Submittal>(LIST_NAMES.Submittals, {
    top: 500,
  });

  const relatedSubmittals = useMemo(() => {
    if (!allSubmittals || !id) return [];
    return allSubmittals.filter((s) => String(s.fields.PropertyLookupId) === String(id));
  }, [allSubmittals, id]);

  // Reset draft when underlying property changes (and we're not in the middle of editing)
  useEffect(() => {
    if (property && !editing) {
      setDraft({ ...property.fields });
    }
  }, [property?.id, property?.lastModifiedDateTime, editing]);

  const handleEdit = () => {
    if (!property) return;
    setDraft({ ...property.fields });
    setSaveError(null);
    setEditing(true);
  };

  const handleCancel = () => {
    if (property) setDraft({ ...property.fields });
    setSaveError(null);
    setEditing(false);
  };

  const handleFieldChange = <K extends keyof PropertyFields>(
    field: K, value: PropertyFields[K]
  ) => {
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
  };

  const handleSave = async () => {
    if (!draft || !property) return;
    setSaving(true);
    setSaveError(null);
    try {
      const changes: Record<string, unknown> = {};
      (Object.keys(draft) as (keyof PropertyFields)[]).forEach((k) => {
        const oldVal = property.fields[k];
        const newVal = draft[k];
        if (oldVal !== newVal) {
          changes[k as string] = newVal === '' ? null : newVal;
        }
      });
      if (Object.keys(changes).length > 0) {
        await updateListItem(LIST_NAMES.Properties, property.id, changes);
        await refetch();
      }
      setEditing(false);
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return (
      <div className="text-center py-12 text-gray-500">
        <div className="inline-flex items-center gap-3">
          <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
          <span className="text-sm">Loading property…</span>
        </div>
      </div>
    );
  }

  if (error || !property || !draft) {
    return (
      <div>
        <BreadcrumbBar />
        <div className="bg-red-50 border border-red-200 rounded-lg p-6">
          <div className="font-semibold text-error mb-2 flex items-center gap-2">
            <Icon name="alert" size={18} />
            Property not found
          </div>
          <p className="text-sm text-red-700 mb-3">
            {error ? error.message : `No property with ID ${id} exists in Properties Registry.`}
          </p>
          <button
            onClick={() => navigate('/properties')}
            className="text-sm text-teal-700 hover:text-teal-900 font-medium underline"
          >
            ← Back to Properties
          </button>
        </div>
      </div>
    );
  }

  const display = editing ? draft : property.fields;

  return (
    <div>
      <BreadcrumbBar propertyName={property.fields.Title} />

      <div className="mb-6 flex items-start justify-between flex-wrap gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-3 flex-wrap">
            <h1 className="text-3xl font-bold text-teal-700">{property.fields.Title}</h1>
            {property.fields.PropertyStatus && (
              <span className={`inline-block px-2.5 py-0.5 rounded text-xs font-semibold border ${STATUS_STYLES[property.fields.PropertyStatus]}`}>
                {property.fields.PropertyStatus}
              </span>
            )}
            {property.fields.cahpVerificationStatus === 'Inherited - Unverified' && (
              <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-yellow-50 text-yellow-800 border border-yellow-200">
                UNVERIFIED
              </span>
            )}
            {property.fields.cahpVerificationStatus === 'Needs Follow-Up' && (
              <span className="inline-block px-2 py-0.5 rounded text-[11px] font-semibold bg-error/10 text-error border border-error/20">
                NEEDS FOLLOW-UP
              </span>
            )}
          </div>
          <p className="text-sm text-gray-500 mt-1">
            {[
              property.fields.LegalEntity,
              property.fields.cahpState,
              property.fields.cahpCounty?.replace(/\s*\([^)]*\)\s*/g, ''),
              property.fields.UnitCount ? `${property.fields.UnitCount} units` : null,
              property.fields.AMIProgram && property.fields.AMIProgram !== 'None'
                ? `${property.fields.AMIProgram} AMI` : null,
            ].filter(Boolean).join(' · ')}
          </p>
          {property.fields.PropertyEIN && (
            <p className="text-xs text-gray-600 mt-0.5 font-mono-data">
              EIN: <span className="select-all">{property.fields.PropertyEIN}</span>
              {property.fields.DORAccountID && (
                <span className="ml-3">DOR Account: <span className="select-all">{property.fields.DORAccountID}</span></span>
              )}
            </p>
          )}
        </div>
        <div className="flex items-center gap-2">
          {!editing && (
            <>
              <button
                onClick={() => setChecklistOpen(true)}
                className="px-3 py-1.5 border border-gold-300 bg-gold-50 hover:bg-gold-100 text-teal-900 rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors"
                title="Generate the DOR-aligned 12-item filing checklist for this property"
              >
                <Icon name="file" size={14} />
                Filing Checklist
              </button>
              {(property.fields.PropertyStatus === 'Active' || property.fields.PropertyStatus === 'Pending') && (
                <button
                  onClick={() => setDispositionOpen(true)}
                  className="px-3 py-1.5 border border-red-300 text-error hover:bg-red-50 rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors"
                  title="Dispose this property (sold, withdrawn, or removed)"
                >
                  <Icon name="alert" size={14} />
                  Dispose
                </button>
              )}
              {isAdmin && (
                <button
                  onClick={() => setDeleteOpen(true)}
                  className="px-3 py-1.5 border-2 border-red-400 bg-red-50 text-error hover:bg-red-100 hover:border-red-500 rounded-md text-sm font-bold flex items-center gap-1.5 transition-colors"
                  title="Permanently delete this property and all related records (Admin only)"
                >
                  <Icon name="alert" size={14} />
                  Delete
                </button>
              )}
              <button
                onClick={handleEdit}
                className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors"
              >
                <Icon name="settings" size={14} />
                Edit
              </button>
            </>
          )}
          {editing && (
            <>
              <button
                onClick={handleCancel} disabled={saving}
                className="px-3 py-1.5 border border-gray-300 text-gray-700 hover:bg-gray-50 rounded-md text-sm font-medium disabled:opacity-50"
              >
                Cancel
              </button>
              <button
                onClick={handleSave} disabled={saving}
                className="px-4 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium flex items-center gap-1.5 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {saving && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
                {saving ? 'Saving…' : 'Save Changes'}
              </button>
            </>
          )}
        </div>
      </div>

      {saveError && (
        <div className="mb-4 bg-red-50 border border-red-200 rounded-lg p-3 flex items-start gap-2">
          <Icon name="alert" size={16} className="text-error flex-shrink-0 mt-0.5" />
          <div className="flex-1 text-sm">
            <span className="font-semibold text-error">Save failed.</span>{' '}
            <span className="text-red-700 font-mono-data text-xs">{saveError}</span>
          </div>
        </div>
      )}

      <div className="border-b border-gray-200 mb-6">
        <div className="flex gap-1">
          {TABS.map((t) => {
            const active = activeTab === t.id;
            const disabled = !t.shipped || editing;
            return (
              <button
                key={t.id}
                onClick={() => !disabled && setActiveTab(t.id)}
                disabled={disabled}
                className={`px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
                  active ? 'border-teal-700 text-teal-700'
                    : !disabled ? 'border-transparent text-gray-600 hover:text-teal-700 hover:border-gray-300'
                    : 'border-transparent text-gray-400 cursor-not-allowed'
                }`}
                title={editing ? 'Save or cancel to switch tabs' : !t.shipped ? 'Coming in a future PR' : undefined}
              >
                {t.label}
                {!t.shipped && <span className="ml-1.5 text-[10px] text-gray-400">soon</span>}
              </button>
            );
          })}
        </div>
      </div>

      {activeTab === 'overview' && (
        <>
          <OverviewTab display={display} editing={editing} onChange={handleFieldChange} />
          {id && <TaxMapIDsSection propertyId={id} propertyTitle={property.fields.Title} />}
          {id && <DeedsSection propertyId={id} propertyTitle={property.fields.Title} />}
        </>
      )}
      {activeTab === 'submittals' && <SubmittalsTab submittals={relatedSubmittals} property={property} onRefetch={refetch} />}
      {activeTab === 'compliance' && id && <PropertyComplianceTab propertyId={id} />}
      {activeTab === 'ownership' && id && <PropertyOwnershipTab propertyId={id} propertyTitle={property.fields.Title} />}
      {activeTab === 'orgChart' && id && <PropertyOrgChartTab propertyId={id} property={property} />}
      {activeTab === 'correspondence' && id && <PropertyCorrespondenceTab propertyId={id} />}
      {activeTab === 'outstanding' && id && <PropertyOutstandingTab propertyId={id} propertyTitle={property.fields.Title} />}
      {activeTab === 'billing' && id && <PropertyBillingTab propertyId={id} />}
      {activeTab === 'documents' && id && <PropertyDocumentsTab propertyId={id} propertyTitle={property.fields.Title} propertyState={property.fields.cahpState} />}
      {activeTab === 'activity' && id && <PropertyActivityTab propertyId={id} />}
      {activeTab === 'notes' && id && <PropertyNotesTab propertyId={id} propertyTitle={property.fields.Title} />}

      {dispositionOpen && (
        <DispositionModal
          property={property}
          onClose={() => setDispositionOpen(false)}
          onComplete={refetch}
        />
      )}

      {deleteOpen && id && (
        <DeletePropertyModal
          propertyId={id}
          propertyTitle={property.fields.Title}
          onClose={() => setDeleteOpen(false)}
        />
      )}

      {checklistOpen && id && (
        <FilingChecklistGenerator
          propertyId={id}
          propertyTitle={property.fields.Title}
          onClose={() => setChecklistOpen(false)}
          onSuccess={(created, matched) => {
            setChecklistOpen(false);
            setChecklistResult({ created, matched });
          }}
        />
      )}

      {checklistResult && (
        <div className="fixed bottom-4 right-4 z-50 max-w-sm bg-green-50 border border-green-200 rounded-md p-3 shadow-lg flex items-start gap-2">
          <Icon name="check" size={14} className="text-success flex-shrink-0 mt-0.5" />
          <div className="flex-1 text-xs">
            <p className="font-semibold text-green-900">Filing Checklist generated</p>
            <p className="text-green-800 mt-0.5">
              <strong>{checklistResult.created}</strong> items created · <strong>{checklistResult.matched}</strong> auto-matched.
              See the Outstanding tab.
            </p>
          </div>
          <button onClick={() => setChecklistResult(null)} className="text-xs text-green-900 hover:underline">×</button>
        </div>
      )}
    </div>
  );
}

function OverviewTab({
  display, editing, onChange,
}: {
  display: PropertyFields;
  editing: boolean;
  onChange: <K extends keyof PropertyFields>(field: K, value: PropertyFields[K]) => void;
}) {
  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <Section title="Property Identity">
        <EditableField label="Property Name" value={display.Title} editing={editing} required onChange={(v) => onChange('Title', v as string)} />
        <EditableField label="Legal Entity" value={display.LegalEntity} editing={editing} onChange={(v) => onChange('LegalEntity', v as string)} />
        <EditableField label="Property Address" value={display.PropertyAddress} editing={editing} onChange={(v) => onChange('PropertyAddress', v as string)} />
        <EditableField label="DOR Account ID" value={display.DORAccountID} editing={editing} mono onChange={(v) => onChange('DORAccountID', v as string)} />
        <EditableField label="EIN" value={display.PropertyEIN} editing={editing} mono onChange={(v) => onChange('PropertyEIN', v as string)} />
        <EditableField label="Date Added to CAHP" value={display.DateAddedToCAHP} editing={editing} type="date" onChange={(v) => onChange('DateAddedToCAHP', v as string)} />
      </Section>

      <Section title="Affordability">
        <EditableField label="AMI Program" value={display.AMIProgram} editing={editing} type="choice" choices={CHOICES.AMIProgram} onChange={(v) => onChange('AMIProgram', v as PropertyFields['AMIProgram'])} />
        <EditableField label="CAHP Language Added" value={display.CAHPLanguageAdded} editing={editing} type="choice" choices={CHOICES.CAHPLanguageAdded} onChange={(v) => onChange('CAHPLanguageAdded', v as PropertyFields['CAHPLanguageAdded'])} />
        <EditableField label="LURA Executed" value={display.LURAExecuted} editing={editing} type="choice" choices={CHOICES.LURAExecuted} onChange={(v) => onChange('LURAExecuted', v as PropertyFields['LURAExecuted'])} />
        <EditableField label="Operating Agreement Version" value={display.OpAgreementVersion} editing={editing} mono onChange={(v) => onChange('OpAgreementVersion', v as string)} />
      </Section>

      <Section title="Status & Verification">
        <EditableField label="Property Status" value={display.PropertyStatus} editing={editing} type="choice" choices={CHOICES.PropertyStatus} onChange={(v) => onChange('PropertyStatus', v as PropertyFields['PropertyStatus'])} />
        <EditableField label="Verification Status" value={display.cahpVerificationStatus} editing={editing} type="choice" choices={CHOICES.cahpVerificationStatus} onChange={(v) => onChange('cahpVerificationStatus', v as PropertyFields['cahpVerificationStatus'])} />
        <EditableField label="Removed Reason" value={display.RemovedReason} editing={editing} type="textarea" onChange={(v) => onChange('RemovedReason', v as string)} />
      </Section>

      <Section title="Location & Ownership">
        <EditableField label="State" value={display.cahpState} editing={editing} type="choice" choices={CHOICES.cahpState} mono onChange={(v) => onChange('cahpState', v as PropertyFields['cahpState'])} />
        <EditableField label="County" value={display.cahpCounty} editing={editing} type="choice" choices={CHOICES.cahpCounty} onChange={(v) => onChange('cahpCounty', v as string)} />
        <EditableField label="Units" value={display.UnitCount} editing={editing} type="number" mono onChange={(v) => onChange('UnitCount', v as number)} />
      </Section>

      <Section title="Owner Contact" fullWidth>
        <OwnerContactField
          value={display.PropertyOwnerContactLookupId}
          editing={editing}
          onChange={(v) => onChange('PropertyOwnerContactLookupId', v)}
        />
      </Section>

      <Section title="Notes" fullWidth>
        <EditableField
          label="Property Notes" value={display.PropertyNotes} editing={editing}
          type="textarea" rows={5} hideLabel
          onChange={(v) => onChange('PropertyNotes', v as string)}
        />
      </Section>
    </div>
  );
}

/**
 * Owner Contact field for the Property Overview.
 * - Always shows the currently linked contact (read-only display) below the picker.
 * - When the page is in edit mode, the ContactPicker is interactive so you can
 *   pick another contact OR create a new one inline.
 */
function OwnerContactField({
  value,
  editing,
  onChange,
}: {
  value: string | undefined;
  editing: boolean;
  onChange: (v: string | undefined) => void;
}) {
  const contactsList = useSharePointList<Contact>(LIST_NAMES.Contacts, { top: 500 });
  const selected = value
    ? contactsList.data?.find((c) => String(c.id) === String(value))
    : undefined;

  if (editing) {
    return (
      <div>
        <ContactPicker
          value={value}
          onChange={(v) => onChange(v)}
        />
        <p className="text-[11px] text-gray-500 mt-1">
          The primary person to ping about this property. Pick an existing contact or use <strong>Add new contact</strong> in the dropdown.
        </p>
      </div>
    );
  }

  // Read-only display
  if (!selected) {
    return <p className="text-sm text-gray-500 italic">No contact set. Click Edit to add one.</p>;
  }
  return (
    <div className="text-sm">
      <div className="font-medium text-gray-900">{selected.fields.Title}</div>
      <div className="text-xs text-gray-600 mt-0.5 space-x-2">
        {selected.fields.ContactEmail && (
          <a href={`mailto:${selected.fields.ContactEmail}`} className="text-teal-700 hover:text-teal-900 underline">
            {selected.fields.ContactEmail}
          </a>
        )}
        {selected.fields.ContactPhone && <span className="font-mono-data">· {selected.fields.ContactPhone}</span>}
        {selected.fields.ContactRole && <span>· {selected.fields.ContactRole}</span>}
      </div>
    </div>
  );
}

function SubmittalsTab({
  submittals,
  property,
  onRefetch,
}: {
  submittals: Submittal[];
  property: Property;
  onRefetch?: () => void;
}) {
  const navigate = useNavigate();
  const [newOpen, setNewOpen] = useState(false);
  const [bulkOpen, setBulkOpen] = useState(false);
  const taxMapIDs = useSharePointList<TaxMapID>(LIST_NAMES.TaxMapIDs, { top: 1000 });
  const propertyParcels = (taxMapIDs.data ?? []).filter(
    (t) => String(t.fields.LinkedPropertyLookupId ?? '') === String(property.id)
  );

  return (
    <>
      {/* Action bar */}
      <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
        <div className="text-xs text-gray-600">
          {submittals.length} submittal{submittals.length === 1 ? '' : 's'} · {propertyParcels.length} tax map ID{propertyParcels.length === 1 ? '' : 's'} registered
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => setNewOpen(true)}
            className="border border-teal-700 text-teal-700 hover:bg-teal-50 px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
          >
            <Icon name="plus" size={12} />
            New Submittal
          </button>
          {propertyParcels.length > 1 && (
            <button
              onClick={() => setBulkOpen(true)}
              className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
            >
              <Icon name="plus" size={12} />
              Bulk create for all {propertyParcels.length} parcels
            </button>
          )}
        </div>
      </div>

      {submittals.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500 mb-3">No submittals on file for this property yet.</p>
          {propertyParcels.length === 0 ? (
            <p className="text-xs text-amber-700 italic">
              Add tax map IDs on the Overview tab first, then create submittals here.
            </p>
          ) : (
            <p className="text-xs text-gray-600">
              Use "New Submittal" for a single submittal, or "Bulk create" to generate one per tax map ID at once.
            </p>
          )}
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Submittal</th>
                <th className="px-4 py-3 text-left">Tax Map ID</th>
                <th className="px-4 py-3 text-left">Tax Year</th>
                <th className="px-4 py-3 text-left">Filing Type</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Filed</th>
                <th className="px-4 py-3 text-right">Abatement</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {submittals.map((s) => {
                const parcel = s.fields.TaxMapIDLookupId
                  ? propertyParcels.find((p) => String(p.id) === String(s.fields.TaxMapIDLookupId))
                  : null;
                return (
                  <tr
                    key={s.id}
                    onClick={() => navigate(`/submittals/${s.id}`)}
                    className="hover:bg-gray-50 transition-colors cursor-pointer"
                  >
                    <td className="px-4 py-3 font-medium text-gray-900">{s.fields.Title}</td>
                    <td className="px-4 py-3 font-mono-data text-xs text-gray-700">
                      {parcel ? parcel.fields.Title : <span className="text-gray-400 italic font-sans">unassigned</span>}
                    </td>
                    <td className="px-4 py-3 text-gray-700 font-mono-data">{s.fields.cahpTaxYear || '—'}</td>
                    <td className="px-4 py-3 text-gray-700 text-xs">{s.fields.FilingType || '—'}</td>
                    <td className="px-4 py-3">
                      {s.fields.SubmittalStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${SUBMITTAL_STATUS_STYLES[s.fields.SubmittalStatus] || 'bg-gray-100'}`}>
                          {s.fields.SubmittalStatus}
                        </span>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3 text-gray-700 text-xs">{formatDate(s.fields.DateFiled)}</td>
                    <td className="px-4 py-3 text-right font-mono-data text-xs">
                      {s.fields.ApprovedAbatement ? `$${s.fields.ApprovedAbatement.toLocaleString()}` : '—'}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {newOpen && (
        <NewSubmittalModal
          fixedPropertyId={String(property.id)}
          onClose={() => setNewOpen(false)}
          onCreated={(id) => {
            onRefetch?.();
            navigate(`/submittals/${id}`);
          }}
        />
      )}

      {bulkOpen && (
        <BulkCreateSubmittalsModal
          propertyId={String(property.id)}
          propertyTitle={property.fields.Title}
          propertyState={property.fields.cahpState}
          onClose={() => setBulkOpen(false)}
          onCreated={() => {
            onRefetch?.();
          }}
        />
      )}
    </>
  );
}

// =============================================================================
// Tab: Compliance — this property's deadlines
// =============================================================================

const DEADLINE_STATUS_STYLES: Record<DeadlineStatus, string> = {
  Upcoming: 'bg-blue-100 text-blue-800',
  'In Progress': 'bg-amber-100 text-amber-800',
  Completed: 'bg-green-100 text-green-800',
  Overdue: 'bg-red-100 text-red-800',
  Missed: 'bg-red-200 text-red-900',
};

function PropertyComplianceTab({ propertyId }: { propertyId: string }) {
  const navigate = useNavigate();
  const { data: allDeadlines, loading, error } = useSharePointList<ComplianceDeadline>(
    LIST_NAMES.ComplianceDeadlines,
    { top: 500 }
  );

  const propertyDeadlines = useMemo(() => {
    if (!allDeadlines) return [];
    return allDeadlines
      .filter((d) => String(d.fields.PropertyLookupId) === String(propertyId))
      .sort((a, b) => {
        const da = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Number.MAX_VALUE;
        const db = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Number.MAX_VALUE;
        return da - db;
      });
  }, [allDeadlines, propertyId]);

  if (loading) return <TabLoading label="deadlines" />;
  if (error) return <TabError error={error} />;

  if (propertyDeadlines.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
        <p className="text-sm text-gray-500">
          No compliance deadlines tied to this property.
        </p>
        <p className="text-xs text-gray-400 mt-2">
          To link a deadline to this property, open Compliance and set the Property field on that deadline.
        </p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <table className="w-full text-sm">
        <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
          <tr>
            <th className="px-4 py-3 text-left">Deadline</th>
            <th className="px-4 py-3 text-left">Type</th>
            <th className="px-4 py-3 text-left">Due Date</th>
            <th className="px-4 py-3 text-left">Status</th>
            <th className="px-4 py-3 text-left">Owner</th>
            <th className="px-4 py-3 text-left">Recurrence</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {propertyDeadlines.map((d) => (
            <tr
              key={d.id}
              onClick={() => navigate(`/compliance/${d.id}`)}
              className="hover:bg-gray-50 transition-colors cursor-pointer"
            >
              <td className="px-4 py-3 font-medium text-gray-900">{d.fields.Title}</td>
              <td className="px-4 py-3 text-gray-700 text-xs">{d.fields.DeadlineType || '—'}</td>
              <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                {formatDate(d.fields.DueDate) || '—'}
              </td>
              <td className="px-4 py-3">
                {d.fields.DeadlineStatus ? (
                  <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${DEADLINE_STATUS_STYLES[d.fields.DeadlineStatus] || 'bg-gray-100'}`}>
                    {d.fields.DeadlineStatus}
                  </span>
                ) : '—'}
              </td>
              <td className="px-4 py-3 text-gray-700">{d.fields.ResponsibleParty || '—'}</td>
              <td className="px-4 py-3 text-gray-700 text-xs">{d.fields.Recurrence || '—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// =============================================================================
// Tab: Correspondence — DOR letters for this property (PR-09c)
// =============================================================================

function PropertyCorrespondenceTab({ propertyId }: { propertyId: string }) {
  const navigate = useNavigate();
  const [logModalOpen, setLogModalOpen] = useState(false);
  const { data, loading, error, refetch } = useSharePointList<Correspondence>(
    LIST_NAMES.Correspondence,
    { top: 500 }
  );

  const filtered = useMemo(() => {
    if (!data) return [];
    return data
      .filter((c) => String(c.fields.PropertyLookupId) === String(propertyId))
      .sort((a, b) => {
        const da = a.fields.DateReceived ? new Date(a.fields.DateReceived).getTime() : 0;
        const db = b.fields.DateReceived ? new Date(b.fields.DateReceived).getTime() : 0;
        return db - da;
      });
  }, [data, propertyId]);

  if (loading) return <TabLoading label="DOR correspondence" />;
  if (error) return <TabError error={error} />;

  const handleLogSuccess = () => {
    setLogModalOpen(false);
    refetch?.();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-gray-700">
          {filtered.length === 0
            ? 'No DOR correspondence on file'
            : `${filtered.length} letter${filtered.length === 1 ? '' : 's'} for this property`}
        </h3>
        <button
          onClick={() => setLogModalOpen(true)}
          className="text-xs text-teal-700 hover:text-teal-900 font-medium flex items-center gap-1"
        >
          <Icon name="plus" size={12} />
          Log Letter
        </button>
      </div>

      {filtered.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500 mb-3">No DOR correspondence tied to this property yet.</p>
          <button
            onClick={() => setLogModalOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
          >
            <Icon name="plus" size={14} />
            Log First Letter
          </button>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Subject</th>
                <th className="px-4 py-3 text-left">Direction</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Received</th>
                <th className="px-4 py-3 text-left">Response Due</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((c) => (
                <tr
                  key={c.id}
                  onClick={() => navigate(`/correspondence/${c.id}`)}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">{c.fields.Title}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{c.fields.Direction || '—'}</td>
                  <td className="px-4 py-3 text-xs text-gray-700">{c.fields.LetterType || '—'}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{formatDate(c.fields.DateReceived)}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{formatDate(c.fields.ResponseDue)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {logModalOpen && (
        <LogLetterModal
          onClose={() => setLogModalOpen(false)}
          onSuccess={handleLogSuccess}
          defaultPropertyId={propertyId}
        />
      )}
    </div>
  );
}

// =============================================================================
// Tab: Outstanding Items — task list scoped to this property
// =============================================================================

const ITEM_STATUS_STYLES: Record<ItemStatus, string> = {
  'Not Started': 'bg-gray-100 text-gray-800',
  'In Progress': 'bg-blue-100 text-blue-800',
  'Blocked': 'bg-red-100 text-red-800',
  'Done': 'bg-green-100 text-green-800',
  'Requested': 'bg-gray-100 text-gray-800',
  'Overdue': 'bg-amber-100 text-amber-800',
  'Received': 'bg-green-100 text-green-800',
  'Not Applicable': 'bg-gray-100 text-gray-500',
};

const ITEM_PRIORITY_STYLES: Record<ItemPriority, string> = {
  Critical: 'bg-red-100 text-red-800',
  High: 'bg-amber-100 text-amber-800',
  Medium: 'bg-blue-100 text-blue-800',
  Low: 'bg-gray-100 text-gray-600',
};

function PropertyOutstandingTab({ propertyId, propertyTitle }: { propertyId: string; propertyTitle: string }) {
  const navigate = useNavigate();
  const [newItemOpen, setNewItemOpen] = useState(false);
  const [showClosed, setShowClosed] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [linkUploadItem, setLinkUploadItem] = useState<OutstandingItem | null>(null);
  const { data, loading, error, refetch } = useSharePointList<OutstandingItem>(
    LIST_NAMES.Outstanding,
    { top: 500 }
  );

  const isClosed = (s: string | undefined) =>
    s === 'Done' || s === 'Received' || s === 'Not Applicable';

  const filtered = useMemo(() => {
    if (!data) return [];
    return data
      .filter((o) => {
        if (String(o.fields.PropertyLookupId) !== String(propertyId)) return false;
        if (!showClosed && isClosed(o.fields.ItemStatus)) return false;
        return true;
      })
      .sort((a, b) => {
        const aOverdue =
          a.fields.DueDate &&
          new Date(a.fields.DueDate).getTime() < Date.now() &&
          !isClosed(a.fields.ItemStatus);
        const bOverdue =
          b.fields.DueDate &&
          new Date(b.fields.DueDate).getTime() < Date.now() &&
          !isClosed(b.fields.ItemStatus);
        if (aOverdue && !bOverdue) return -1;
        if (!aOverdue && bOverdue) return 1;
        const aDue = a.fields.DueDate ? new Date(a.fields.DueDate).getTime() : Infinity;
        const bDue = b.fields.DueDate ? new Date(b.fields.DueDate).getTime() : Infinity;
        return aDue - bDue;
      });
  }, [data, propertyId, showClosed]);

  const totalForProperty = useMemo(() => {
    if (!data) return 0;
    return data.filter((o) => String(o.fields.PropertyLookupId) === String(propertyId)).length;
  }, [data, propertyId]);

  if (loading) return <TabLoading label="outstanding items" />;
  if (error) return <TabError error={error} />;

  return (
    <div>
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <h3 className="text-sm font-semibold text-gray-700">
          {filtered.length === 0
            ? showClosed
              ? 'No items on file'
              : 'No open items'
            : `${filtered.length} ${showClosed ? 'item' : 'open item'}${filtered.length === 1 ? '' : 's'} for this property`}
        </h3>
        <div className="flex items-center gap-2">
          <label className="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer">
            <input
              type="checkbox"
              checked={showClosed}
              onChange={(e) => setShowClosed(e.target.checked)}
            />
            Show closed
          </label>
          <button
            onClick={() => setExportOpen(true)}
            disabled={filtered.filter((i) => !isClosed(i.fields.ItemStatus)).length === 0}
            className="bg-white border border-gray-300 hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed text-gray-700 px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
            title="Export a copy-pastable list of open items for an assignee"
          >
            <Icon name="file" size={12} />
            Export
          </button>
          <button
            onClick={() => setNewItemOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5"
          >
            <Icon name="plus" size={12} />
            Add Item
          </button>
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500 mb-3">
            {totalForProperty === 0
              ? 'No outstanding items tied to this property yet.'
              : showClosed
                ? 'No items match the current filter.'
                : 'All items for this property are closed.'}
          </p>
          <button
            onClick={() => setNewItemOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
          >
            <Icon name="plus" size={14} />
            Add First Item
          </button>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Title</th>
                <th className="px-4 py-3 text-left">Status</th>
                <th className="px-4 py-3 text-left">Priority</th>
                <th className="px-4 py-3 text-left">Due Date</th>
                <th className="px-4 py-3 text-left">Category</th>
                <th className="px-4 py-3 text-left">Document</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map((o) => {
                const overdue =
                  o.fields.DueDate &&
                  new Date(o.fields.DueDate).getTime() < Date.now() &&
                  !isClosed(o.fields.ItemStatus);
                const hasDoc = Boolean(o.fields.RelatedDocUrl);
                return (
                  <tr
                    key={o.id}
                    className={`hover:bg-gray-50 transition-colors ${overdue ? 'bg-red-50' : ''}`}
                  >
                    <td
                      className="px-4 py-3 font-medium text-gray-900 cursor-pointer"
                      onClick={() => navigate(`/outstanding-items/${o.id}`)}
                    >
                      {overdue && <span className="text-error mr-1">⚠</span>}
                      {o.fields.Title}
                    </td>
                    <td
                      className="px-4 py-3 cursor-pointer"
                      onClick={() => navigate(`/outstanding-items/${o.id}`)}
                    >
                      {o.fields.ItemStatus ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${ITEM_STATUS_STYLES[o.fields.ItemStatus]}`}>
                          {o.fields.ItemStatus}
                        </span>
                      ) : '—'}
                    </td>
                    <td
                      className="px-4 py-3 cursor-pointer"
                      onClick={() => navigate(`/outstanding-items/${o.id}`)}
                    >
                      {o.fields.Priority ? (
                        <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${ITEM_PRIORITY_STYLES[o.fields.Priority]}`}>
                          {o.fields.Priority}
                        </span>
                      ) : '—'}
                    </td>
                    <td
                      className={`px-4 py-3 font-mono-data text-xs cursor-pointer ${overdue ? 'text-error font-semibold' : 'text-gray-700'}`}
                      onClick={() => navigate(`/outstanding-items/${o.id}`)}
                    >
                      {formatDateOnly(o.fields.DueDate)}
                    </td>
                    <td
                      className="px-4 py-3 text-xs text-gray-600 cursor-pointer"
                      onClick={() => navigate(`/outstanding-items/${o.id}`)}
                    >
                      {o.fields.ItemCategory || '—'}
                    </td>
                    <td className="px-4 py-3 text-xs">
                      {hasDoc ? (
                        <a
                          href={o.fields.RelatedDocUrl!}
                          target="_blank"
                          rel="noopener noreferrer"
                          onClick={(e) => e.stopPropagation()}
                          className="text-success hover:underline inline-flex items-center gap-1 max-w-[180px] truncate"
                          title={o.fields.RelatedDocFilename}
                        >
                          <Icon name="check" size={11} />
                          <span className="truncate">{o.fields.RelatedDocFilename ?? 'View'}</span>
                        </a>
                      ) : (
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setLinkUploadItem(o);
                          }}
                          className="text-gold-700 hover:text-gold-900 underline inline-flex items-center gap-1"
                        >
                          <Icon name="plus" size={11} />
                          Link / Upload
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {newItemOpen && (
        <NewOutstandingItemModal
          onClose={() => setNewItemOpen(false)}
          onSuccess={() => {
            setNewItemOpen(false);
            refetch?.();
          }}
          defaultPropertyId={propertyId}
          hidePropertyPicker
        />
      )}

      {linkUploadItem && (
        <LinkOrUploadDocumentModal
          item={linkUploadItem}
          onClose={() => setLinkUploadItem(null)}
          onSuccess={() => {
            setLinkUploadItem(null);
            refetch?.();
          }}
        />
      )}

      {exportOpen && (
        <ExportOutstandingItemsModal
          items={filtered.filter((i) => !isClosed(i.fields.ItemStatus))}
          propertyTitle={propertyTitle}
          onClose={() => setExportOpen(false)}
        />
      )}
      {/* propertyTitle ref to avoid TS unused warning if banner removed later */}
      <span className="hidden" data-property-title={propertyTitle} />
    </div>
  );
}

// =============================================================================
// Tab: Billing — fee invoices for this property (PR-09c)
// =============================================================================

const BILLING_STATUS_STYLES: Record<string, string> = {
  'Pending Approval': 'bg-amber-100 text-amber-800',
  'Ready to Invoice': 'bg-blue-100 text-blue-800',
  'Invoiced': 'bg-indigo-100 text-indigo-800',
  'Paid': 'bg-green-100 text-green-800',
  'Disputed': 'bg-red-100 text-red-800',
};

function PropertyBillingTab({ propertyId }: { propertyId: string }) {
  const { data, loading, error } = useSharePointList<Billing>(
    LIST_NAMES.Billing,
    { top: 500 }
  );

  const filtered = useMemo(() => {
    if (!data) return [];
    return data
      .filter((b) => String(b.fields.PropertyLookupId) === String(propertyId))
      .sort((a, b) => {
        const ya = Number(a.fields.cahpTaxYear ?? 0);
        const yb = Number(b.fields.cahpTaxYear ?? 0);
        return yb - ya;
      });
  }, [data, propertyId]);

  if (loading) return <TabLoading label="billing records" />;
  if (error) return <TabError error={error} />;

  if (filtered.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
        <p className="text-sm text-gray-500">No billing records for this property yet.</p>
        <p className="text-xs text-gray-400 mt-2">
          The Billing module is intentionally deferred. Status tracking on submittals (Approved, ApprovedAbatement)
          continues to work, but no automated billing records are created.
        </p>
      </div>
    );
  }

  const totalBilled = filtered.reduce((sum, b) => sum + (b.fields.AmountBilled ?? 0), 0);
  const totalAbatement = filtered.reduce((sum, b) => sum + (b.fields.BillApprovedAbatement ?? 0), 0);

  return (
    <div>
      <div className="grid grid-cols-2 gap-4 mb-4">
        <div className="bg-white border border-gray-200 rounded-lg p-3 shadow-card">
          <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Total CAHP Fees Billed</div>
          <div className="text-xl font-bold mt-0.5 text-teal-700 font-mono-data">${totalBilled.toLocaleString()}</div>
        </div>
        <div className="bg-white border border-gray-200 rounded-lg p-3 shadow-card">
          <div className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider">Total Abatement (basis)</div>
          <div className="text-xl font-bold mt-0.5 text-gray-700 font-mono-data">${totalAbatement.toLocaleString()}</div>
        </div>
      </div>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Reference</th>
              <th className="px-4 py-3 text-left">Tax Year</th>
              <th className="px-4 py-3 text-right">Abatement</th>
              <th className="px-4 py-3 text-right">Fee %</th>
              <th className="px-4 py-3 text-right">Billed</th>
              <th className="px-4 py-3 text-left">Status</th>
              <th className="px-4 py-3 text-left">QB Sync</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filtered.map((b) => (
              <tr key={b.id} className="hover:bg-gray-50 transition-colors">
                <td className="px-4 py-3 font-medium text-gray-900">{b.fields.Title}</td>
                <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">{b.fields.cahpTaxYear || '—'}</td>
                <td className="px-4 py-3 text-right font-mono-data text-xs">
                  {b.fields.BillApprovedAbatement ? `$${b.fields.BillApprovedAbatement.toLocaleString()}` : '—'}
                </td>
                <td className="px-4 py-3 text-right font-mono-data text-xs">
                  {b.fields.CAHPFeePercent != null ? `${b.fields.CAHPFeePercent}%` : '—'}
                </td>
                <td className="px-4 py-3 text-right font-mono-data font-semibold">
                  {b.fields.AmountBilled ? `$${b.fields.AmountBilled.toLocaleString()}` : '—'}
                </td>
                <td className="px-4 py-3">
                  {b.fields.BillingStatus ? (
                    <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold ${BILLING_STATUS_STYLES[b.fields.BillingStatus] || 'bg-gray-100'}`}>
                      {b.fields.BillingStatus}
                    </span>
                  ) : '—'}
                </td>
                <td className="px-4 py-3 text-xs text-gray-700">{b.fields.QBSyncStatus || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// =============================================================================
// Tab: Activity — audit log filtered to this property (PR-09c)
// =============================================================================

const ACTIVITY_ACTION_STYLES: Record<string, string> = {
  CREATE: 'bg-green-100 text-green-800',
  UPDATE: 'bg-blue-100 text-blue-800',
  DELETE: 'bg-red-100 text-red-800',
};

function PropertyActivityTab({ propertyId }: { propertyId: string }) {
  const { data, loading, error } = useSharePointList<AuditLog>(
    LIST_NAMES.AuditLog,
    { top: 500 }
  );

  const filtered = useMemo(() => {
    if (!data) return [];
    // Direct property changes
    return data
      .filter(
        (a) =>
          a.fields.EntityType === 'Property' &&
          String(a.fields.EntityId) === String(propertyId)
      )
      .sort(
        (a, b) =>
          new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime()
      );
  }, [data, propertyId]);

  if (loading) return <TabLoading label="activity" />;
  if (error) return <TabError error={error} />;

  if (filtered.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
        <p className="text-sm text-gray-500">No activity recorded for this property yet.</p>
        <p className="text-xs text-gray-400 mt-2">
          Every edit, create, or delete of this property is captured here. For broader audit data
          (changes to related submittals, ownership, etc.) see the full <Link to="/audit" className="text-teal-700 underline">Audit Log</Link>.
        </p>
      </div>
    );
  }

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100">
        <h3 className="text-sm font-semibold text-gray-700">
          {filtered.length} event{filtered.length === 1 ? '' : 's'} affecting this property
        </h3>
        <p className="text-xs text-gray-500 mt-0.5">Newest first · click an event to see full diff in the Audit Log</p>
      </div>
      <ul className="divide-y divide-gray-100">
        {filtered.map((row) => (
          <li key={row.id} className="px-4 py-3 flex items-start gap-3 text-sm">
            <span
              className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-bold flex-shrink-0 ${
                ACTIVITY_ACTION_STYLES[row.fields.Action ?? ''] ?? 'bg-gray-100'
              }`}
            >
              {row.fields.Action}
            </span>
            <div className="flex-1 min-w-0">
              <div className="text-gray-900">{row.fields.Title}</div>
              {row.fields.ChangeSummary && (
                <pre className="mt-1 text-xs text-gray-600 whitespace-pre-wrap font-mono-data">
                  {row.fields.ChangeSummary}
                </pre>
              )}
            </div>
            <div className="text-xs text-gray-500 font-mono-data flex-shrink-0 text-right">
              <div>{new Date(row.createdDateTime).toLocaleDateString()}</div>
              <div className="text-gray-400">{row.createdBy?.user?.displayName ?? ''}</div>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
}

// =============================================================================
// Tab: Org Chart — visual hierarchy with 3 layouts (PR-09b)
// =============================================================================

type ChartLayout = 'detailed' | 'beneficial' | 'dor';

const LAYOUT_INFO: Record<ChartLayout, { label: string; description: string }> = {
  detailed: {
    label: 'Detailed',
    description: 'Full ownership chain — property at top, members below, recursive until natural persons or nonprofits.',
  },
  beneficial: {
    label: 'Beneficial',
    description: 'Collapsed to natural-person and nonprofit beneficial owners with compounded percentages.',
  },
  dor: {
    label: 'DOR-Friendly',
    description: 'Property at the bottom per DOR convention. Same data as Detailed, flipped vertical orientation.',
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

function PropertyOrgChartTab({ propertyId, property }: { propertyId: string; property: Property }) {
  const [layout, setLayout] = useState<ChartLayout>('detailed');
  const propertyTitle = property.fields.Title ?? '(unnamed)';

  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });

  const loading = ownership.loading || owners.loading;
  const error = ownership.error || owners.error;

  const tree = useMemo(() => {
    if (!ownership.data || !owners.data) return [];
    return getBeneficialOwnershipTree('property', propertyId, ownership.data, owners.data);
  }, [ownership.data, owners.data, propertyId]);

  const beneficial = useMemo(() => {
    if (!ownership.data || !owners.data) return [];
    return computeBeneficialOwnership('property', propertyId, ownership.data, owners.data);
  }, [ownership.data, owners.data, propertyId]);

  if (loading) return <TabLoading label="org chart" />;
  if (error) return <TabError error={error} />;

  const hasData = tree.length > 0;

  return (
    <div>
      {/* Layout switcher */}
      <div className="bg-white border border-gray-200 rounded-lg p-3 mb-4 shadow-card">
        <div className="flex flex-wrap gap-2 mb-2">
          {(Object.keys(LAYOUT_INFO) as ChartLayout[]).map((key) => (
            <button
              key={key}
              onClick={() => setLayout(key)}
              className={`px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                layout === key
                  ? 'bg-teal-700 text-white'
                  : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
              }`}
            >
              {LAYOUT_INFO[key].label}
            </button>
          ))}
        </div>
        <p className="text-xs text-gray-500">{LAYOUT_INFO[layout].description}</p>
      </div>

      {/* Empty state */}
      {!hasData ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500 mb-2">
            No ownership records — there's nothing to render yet.
          </p>
          <p className="text-xs text-gray-400">
            Add owners to this property via the Ownership tab, then trace the chain by adding member-of relationships on each Owner's detail page.
          </p>
        </div>
      ) : (
        <>
          {layout === 'detailed' && <DetailedChart tree={tree} propertyTitle={propertyTitle} />}
          {layout === 'beneficial' && <BeneficialChart beneficialOwners={beneficial} propertyTitle={propertyTitle} />}
          {layout === 'dor' && <DORChart tree={tree} property={property} owners={owners.data ?? []} />}
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// Layout 1: Detailed (top-down)
// ─────────────────────────────────────────────────────────────

function DetailedChart({ tree, propertyTitle }: { tree: OwnershipNode[]; propertyTitle: string }) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card p-5">
      <PropertyRootNode title={propertyTitle} />
      <div className="pl-6 ml-3 border-l-2 border-gray-300 mt-2 space-y-2">
        {tree.map((node) => (
          <TreeBranch key={node.relationship.id} node={node} />
        ))}
      </div>
    </div>
  );
}

function TreeBranch({ node }: { node: OwnershipNode }) {
  return (
    <div>
      <EntityCard
        name={node.owner?.fields.Title ?? '(unresolved)'}
        ownerType={node.owner?.fields.OwnerType}
        relationshipType={node.relationship.fields.RelationshipType}
        percent={node.relationship.fields.OwnershipPercent}
      />
      {node.children.length > 0 && (
        <div className="pl-6 ml-3 border-l-2 border-gray-300 mt-2 space-y-2">
          {node.children.map((child) => (
            <TreeBranch key={child.relationship.id} node={child} />
          ))}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// Layout 2: Beneficial (flat list of natural persons / nonprofits)
// ─────────────────────────────────────────────────────────────

function BeneficialChart({
  beneficialOwners,
  propertyTitle,
}: {
  beneficialOwners: BeneficialOwner[];
  propertyTitle: string;
}) {
  const totalKnown = beneficialOwners.reduce((sum, b) => sum + b.beneficialPercent, 0);

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
      <div className="px-4 py-3 border-b border-gray-100">
        <h3 className="text-sm font-semibold text-gray-700">Beneficial owners of {propertyTitle}</h3>
        <p className="text-xs text-gray-500 mt-0.5">
          {beneficialOwners.length} terminal owner{beneficialOwners.length === 1 ? '' : 's'} · {totalKnown.toFixed(2)}% of ownership chain traced
          {totalKnown < 99.9 && (
            <span className="text-warning"> · {(100 - totalKnown).toFixed(2)}% not yet traced (missing upstream relationships)</span>
          )}
        </p>
      </div>
      {beneficialOwners.length === 0 ? (
        <div className="px-4 py-8 text-center text-sm text-gray-500">
          No natural-person or nonprofit terminal owners found. Add upstream member relationships on the Owner detail pages.
        </div>
      ) : (
        <ul className="divide-y divide-gray-100">
          {beneficialOwners.map((b) => (
            <li key={b.owner.id} className="px-4 py-3">
              <div className="flex items-center justify-between gap-3">
                <div className="flex items-center gap-2 flex-1 min-w-0">
                  <span className="font-medium text-gray-900 truncate">{b.owner.fields.Title}</span>
                  {b.owner.fields.OwnerType && (
                    <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold flex-shrink-0 ${OWNER_TYPE_BADGE_STYLES[b.owner.fields.OwnerType]}`}>
                      {b.owner.fields.OwnerType}
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
                <div className="mt-1.5 text-xs text-gray-500">
                  {b.paths.map((path, idx) => (
                    <div key={idx} className="font-mono-data">
                      via {path.intermediates.length === 0 ? '(direct)' : path.intermediates.map((i) => i.owner.fields.Title).join(' ← ')}
                      {' · '}{path.pathPercent.toFixed(2)}%
                    </div>
                  ))}
                </div>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// Layout 3: DOR-Friendly (bottom-up, property at the bottom)
// ─────────────────────────────────────────────────────────────

function DORChart({ tree, property, owners }: { tree: OwnershipNode[]; property: Property; owners: Owner[] }) {
  // Identify the "Manager-Managed by X" line — pick the direct owner whose role is Managing Member
  const managerName = useMemo(() => {
    const managerNode = tree.find(
      (n) => n.relationship.fields.RelationshipType === 'Managing Member'
    );
    if (!managerNode?.owner) return undefined;
    return managerNode.owner.fields.Title;
  }, [tree]);

  // Resolve the legal entity name. Use Property.LegalEntity if set, otherwise fall back to Title.
  const legalEntity = property.fields.LegalEntity || property.fields.Title || '(unnamed)';

  // Suppress unused for the lint; owners is reserved for future enhancements
  void owners;

  // PDF export state
  const orgChartsLibrary = useSharePointList<{ id: string; fields: { FileLeafRef?: string } }>('Org Charts', { top: 500 });
  const [exporting, setExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState<string>('');
  const [exportResult, setExportResult] = useState<{ filename: string; webUrl: string } | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);

  const handleExport = async () => {
    setExporting(true);
    setExportError(null);
    setExportResult(null);
    try {
      const { exportOrgChartPDF } = await import('../components/exportOrgChartPDF');
      const existing = (orgChartsLibrary.data ?? [])
        .map((d) => d.fields.FileLeafRef ?? '')
        .filter(Boolean);
      const result = await exportOrgChartPDF({
        tree,
        property,
        managerName,
        existingFilenames: existing,
        onProgress: (_pct, label) => setExportProgress(label),
      });
      setExportResult(result);
      orgChartsLibrary.refetch?.();
    } catch (err) {
      setExportError(err instanceof Error ? err.message : String(err));
    } finally {
      setExporting(false);
      setExportProgress('');
    }
  };

  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-card p-5 overflow-x-auto">
      <div className="flex items-start justify-between mb-4 gap-3">
        <p className="text-xs text-gray-500 italic flex-1">
          Beneficial owners at the top, property at the bottom — the orientation DOR prefers for org chart submissions.
        </p>
        <button
          onClick={handleExport}
          disabled={exporting || tree.length === 0}
          className="bg-teal-700 hover:bg-teal-900 disabled:bg-gray-300 disabled:cursor-not-allowed text-white px-3 py-1.5 rounded-md text-xs font-medium inline-flex items-center gap-1.5 flex-shrink-0"
        >
          {exporting ? (
            <>
              <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />
              {exportProgress || 'Exporting…'}
            </>
          ) : (
            <>
              <Icon name="file" size={12} />
              Save to Documents
            </>
          )}
        </button>
      </div>

      {exportResult && (
        <div className="mb-3 bg-green-50 border border-green-200 rounded-lg px-3 py-2 text-xs text-green-900 flex items-center justify-between gap-2">
          <span>
            <Icon name="check" size={12} className="inline mr-1" />
            Saved <strong>{exportResult.filename}</strong> to the Org Charts library, tagged to {property.fields.Title}.
          </span>
          <a
            href={exportResult.webUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="text-teal-700 hover:text-teal-900 font-medium underline"
          >
            Open
          </a>
        </div>
      )}

      {exportError && (
        <div className="mb-3 bg-red-50 border border-red-200 rounded-lg px-3 py-2 text-xs text-red-900">
          <Icon name="alert" size={12} className="inline mr-1" />
          Export failed: {exportError}
        </div>
      )}

      <div className="flex flex-col items-center min-w-fit bg-white p-3">
        {/* Top section: one column per direct owner of the property.
            Uses the same T-junction pattern as nested levels — half-bars
            converge at center, single arrow goes down to the property. */}
        <div className="flex flex-row items-end justify-center" style={{ gap: '1rem' }}>
          {tree.map((directOwner, idx) => {
            const isFirst = idx === 0;
            const isLast = idx === tree.length - 1;
            const isOnly = tree.length === 1;
            return (
              <div
                key={directOwner.relationship.id}
                className="relative flex flex-col items-center"
              >
                <DORColumn node={directOwner} />
                {isOnly ? (
                  <div className="w-px h-4 bg-gray-400 mt-1" />
                ) : (
                  <div className="relative w-full h-4 mt-1">
                    <div className="absolute top-0 left-1/2 -ml-px w-px h-2 bg-gray-400" />
                    <div
                      className="absolute h-px bg-gray-400"
                      style={{
                        top: '0.5rem',
                        left: isFirst ? '50%' : '-0.5rem',
                        right: isLast ? '50%' : '-0.5rem',
                      }}
                    />
                  </div>
                )}
              </div>
            );
          })}
        </div>
        {/* Center arrow down to property */}
        {tree.length > 0 && (
          <div className="text-gray-400">
            <svg width="12" height="14" viewBox="0 0 12 14" xmlns="http://www.w3.org/2000/svg">
              <line x1="6" y1="0" x2="6" y2="8" stroke="currentColor" strokeWidth="1" />
              <polygon points="6,14 2,8 10,8" fill="currentColor" />
            </svg>
          </div>
        )}

        {/* Property — full DOR-style card */}
        <div>
          <PropertyRootNode
            legalEntity={legalEntity}
            ownerState={property.fields.cahpState}
            address={property.fields.PropertyAddress}
            managerName={managerName}
          />
        </div>
      </div>
    </div>
  );
}

/**
 * Renders one direct-owner-of-the-property as a vertical column:
 *   - Topmost ancestor at the top
 *   - Each descendant beneath it
 *   - The direct-owner-of-property at the bottom (just above the property row)
 *
 * Walks UP from the direct owner to its terminals via the children array.
 * (Tree shape: direct.children = ancestors of direct, recursively.)
 */
/**
 * Recursive tree renderer for the DOR org chart.
 * Each node renders its parents (children in the tree) ABOVE it, side by side,
 * with arrows converging down. Recurses up the chain until terminal owners.
 *
 * Tree shape reminder: `node.children` are the upstream OWNERS of this node
 * (the relationships where someone owns a piece of this node), NOT subordinates.
 */
function DORColumn({ node, isTopLevel = false }: { node: OwnershipNode; isTopLevel?: boolean }) {
  // "Single-member LLC" inference: exactly one upstream parent owning 100%
  const isSingleMember =
    node.children.length === 1 &&
    node.children[0].relationship.fields.OwnershipPercent === 100;

  void isTopLevel;

  return (
    <div className="flex flex-col items-center gap-0">
      {/* Render all upstream owners ABOVE this node, then a T-junction connector */}
      {node.children.length > 0 && (
        <>
          <div className="flex flex-row items-end justify-center" style={{ gap: '1rem' }}>
            {node.children.map((parent, idx) => {
              const isFirst = idx === 0;
              const isLast = idx === node.children.length - 1;
              const isOnly = node.children.length === 1;
              return (
                <div
                  key={parent.relationship.id}
                  className="relative flex flex-col items-center"
                >
                  <DORColumn node={parent} />
                  {/* Below each parent card: stub + half horizontal bar */}
                  {isOnly ? (
                    <div className="w-px h-4 bg-gray-400 mt-1" />
                  ) : (
                    <div className="relative w-full h-4 mt-1">
                      {/* Vertical stub from card down */}
                      <div className="absolute top-0 left-1/2 -ml-px w-px h-2 bg-gray-400" />
                      {/* Horizontal half-bar at bottom of stub.
                          Extends 0.5rem into the gap on each side so adjacent half-bars meet. */}
                      <div
                        className="absolute h-px bg-gray-400"
                        style={{
                          top: '0.5rem',
                          left: isFirst ? '50%' : '-0.5rem',
                          right: isLast ? '50%' : '-0.5rem',
                        }}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          {/* Single center arrow from the horizontal bar down to this node */}
          <div className="text-gray-400">
            <svg width="12" height="14" viewBox="0 0 12 14" xmlns="http://www.w3.org/2000/svg">
              <line x1="6" y1="0" x2="6" y2="8" stroke="currentColor" strokeWidth="1" />
              <polygon points="6,14 2,8 10,8" fill="currentColor" />
            </svg>
          </div>
        </>
      )}

      {/* This node */}
      <EntityCard
        name={node.owner?.fields.Title ?? '(unresolved)'}
        ownerType={node.owner?.fields.OwnerType}
        relationshipType={node.relationship.fields.RelationshipType}
        percent={node.relationship.fields.OwnershipPercent}
        memberClass={node.relationship.fields.MemberClass}
        sponsorName={node.owner?.fields.SponsorName}
        stateOfFormation={node.owner?.fields.OwnerState}
        isTaxExempt={node.owner?.fields.IsTaxExempt}
        entityDescription={node.owner?.fields.EntityDescription}
        isSingleMember={isSingleMember}
      />
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// Shared node components
// ─────────────────────────────────────────────────────────────

function PropertyRootNode({
  title,
  legalEntity,
  ownerState,
  address,
  managerName,
}: {
  title?: string;          // For Detailed / Beneficial views — just a small label
  legalEntity?: string;    // For DOR view — large card with full entity details
  ownerState?: string;
  address?: string;
  managerName?: string;
}) {
  // Detailed-mode rendering: large card matching DOR convention. The PDF
  // export is now drawn natively with jsPDF (see exportOrgChartPDF.ts), so
  // this is on-screen only — Tailwind classes would be fine here too.
  if (legalEntity) {
    const stateName = ownerState ? spellState(ownerState) : '';
    return (
      <div
        className="inline-block rounded-lg shadow-md min-w-[320px] text-center"
        style={{
          backgroundColor: '#0f766e', // teal-700
          color: '#ffffff',
          padding: '16px 24px',
        }}
      >
        <div
          style={{
            color: '#ffffff',
            fontWeight: 700,
            fontSize: '16px',
            textTransform: 'uppercase',
            letterSpacing: '0.05em',
          }}
        >
          {legalEntity}
        </div>
        {stateName && (
          <div style={{ color: '#ccfbf1', fontSize: '12px', marginTop: '4px' }}>
            {stateName} LLC
          </div>
        )}
        {managerName && (
          <div style={{ color: '#ccfbf1', fontSize: '12px' }}>
            Manager-Managed by {managerName}
          </div>
        )}
        {address && (
          <div style={{ color: '#ccfbf1', fontSize: '12px', marginTop: '2px' }}>
            {address}
          </div>
        )}
      </div>
    );
  }

  // Compact mode (Detailed / Beneficial views)
  return (
    <div className="inline-block bg-gold-50 border-2 border-gold-500 rounded-lg px-4 py-2 shadow-sm">
      <div className="flex items-center gap-2">
        <Icon name="folder" size={14} className="text-gold-700" />
        <span className="font-bold text-teal-900 text-sm">{title ?? '(unnamed)'}</span>
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
  memberClass,
  sponsorName,
  stateOfFormation,
  isTaxExempt,
  entityDescription,
  isSingleMember,
}: {
  name: string;
  ownerType?: string;
  relationshipType?: string;
  percent?: number;
  memberClass?: string;
  sponsorName?: string;
  stateOfFormation?: string;
  isTaxExempt?: boolean;
  entityDescription?: string;
  isSingleMember?: boolean;
}) {
  // Build the formation description line. Priority:
  //   1. Manual override (EntityDescription) — use as-is
  //   2. Auto-derive from OwnerType + state + single-member status
  const stateName = stateOfFormation ? spellState(stateOfFormation) : '';
  let typeLine = entityDescription ?? '';
  if (!typeLine) {
    if (ownerType === 'Nonprofit') {
      typeLine = `${stateName ? stateName + ' ' : ''}Nonprofit Corp.`;
    } else if (ownerType === 'LLC') {
      const singleMember = isSingleMember ? 'Single-Member ' : '';
      typeLine = `${stateName ? stateName + ' ' : ''}${singleMember}LLC`;
    } else if (ownerType === 'Trust') {
      typeLine = 'Trust';
    } else if (ownerType === 'Corporation') {
      typeLine = `${stateName ? stateName + ' ' : ''}Corporation`;
    } else if (ownerType === 'Limited Partnership') {
      typeLine = `${stateName ? stateName + ' ' : ''}Limited Partnership`;
    } else if (ownerType === 'General Partnership') {
      typeLine = `${stateName ? stateName + ' ' : ''}General Partnership`;
    } else if (ownerType) {
      typeLine = ownerType;
    }
  }

  return (
    <div className="inline-block bg-white border border-gray-300 rounded-lg px-3 py-2 shadow-sm min-w-[220px] text-center">
      <div className="flex items-center justify-center gap-2 flex-wrap">
        <span className="font-semibold text-gray-900 text-sm">{name}</span>
        {ownerType && (
          <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold ${OWNER_TYPE_BADGE_STYLES[ownerType] ?? 'bg-gray-100 text-gray-700'}`}>
            {ownerType}
          </span>
        )}
      </div>
      {typeLine && (
        <div className="text-[11px] text-gray-600 mt-0.5">{typeLine}</div>
      )}
      {isTaxExempt && ownerType === 'Nonprofit' && (
        <div className="text-[11px] text-gray-600 italic">IRC § 501(c)(3) Tax-Exempt</div>
      )}
      {sponsorName && (
        <div className="text-[11px] text-gray-600 italic">
          Sponsor: {sponsorName}
        </div>
      )}
      <div className="text-xs text-gray-700 mt-1 font-mono-data">
        {memberClass && <span className="mr-1 font-semibold">{memberClass}</span>}
        {relationshipType ?? 'Member'} · {percent != null ? `${percent}%` : '—'}
      </div>
    </div>
  );
}

// US state code → full name (for org chart entity descriptions)
const STATE_NAMES: Record<string, string> = {
  AL: 'Alabama', AK: 'Alaska', AZ: 'Arizona', AR: 'Arkansas', CA: 'California',
  CO: 'Colorado', CT: 'Connecticut', DE: 'Delaware', FL: 'Florida', GA: 'Georgia',
  HI: 'Hawaii', ID: 'Idaho', IL: 'Illinois', IN: 'Indiana', IA: 'Iowa',
  KS: 'Kansas', KY: 'Kentucky', LA: 'Louisiana', ME: 'Maine', MD: 'Maryland',
  MA: 'Massachusetts', MI: 'Michigan', MN: 'Minnesota', MS: 'Mississippi', MO: 'Missouri',
  MT: 'Montana', NE: 'Nebraska', NV: 'Nevada', NH: 'New Hampshire', NJ: 'New Jersey',
  NM: 'New Mexico', NY: 'New York', NC: 'North Carolina', ND: 'North Dakota', OH: 'Ohio',
  OK: 'Oklahoma', OR: 'Oregon', PA: 'Pennsylvania', RI: 'Rhode Island', SC: 'South Carolina',
  SD: 'South Dakota', TN: 'Tennessee', TX: 'Texas', UT: 'Utah', VT: 'Vermont',
  VA: 'Virginia', WA: 'Washington', WV: 'West Virginia', WI: 'Wisconsin', WY: 'Wyoming',
  DC: 'District of Columbia',
};
function spellState(code: string): string {
  const upper = code.trim().toUpperCase();
  return STATE_NAMES[upper] ?? code;
}

// =============================================================================
// Tab: Ownership — this property's ownership chain
// =============================================================================

const RELATIONSHIP_STYLES: Record<RelationshipType, string> = {
  'Managing Member': 'bg-teal-100 text-teal-800',
  'Sole Member': 'bg-emerald-100 text-emerald-800',
  Member: 'bg-blue-100 text-blue-800',
  Owner: 'bg-purple-100 text-purple-800',
  Subsidiary: 'bg-amber-100 text-amber-800',
  'Beneficial Owner': 'bg-pink-100 text-pink-800',
};

function PropertyOwnershipTab({ propertyId, propertyTitle }: { propertyId: string; propertyTitle: string }) {
  const navigate = useNavigate();
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const [editingOwnershipId, setEditingOwnershipId] = useState<string | null>(null);

  const loading = ownership.loading || owners.loading;
  const error = ownership.error || owners.error;

  const ownersById = useMemo(() => {
    if (!owners.data) return new Map<string, Owner>();
    return new Map(owners.data.map((o) => [String(o.id), o]));
  }, [owners.data]);

  const directOwnership = useMemo(() => {
    if (!ownership.data) return [];
    return ownership.data
      .filter((o) => String(o.fields.LinkedPropertyLookupId) === String(propertyId))
      .sort((a, b) => (b.fields.OwnershipPercent ?? 0) - (a.fields.OwnershipPercent ?? 0));
  }, [ownership.data, propertyId]);

  if (loading) return <TabLoading label="ownership records" />;
  if (error) return <TabError error={error} />;

  if (directOwnership.length === 0) {
    return (
      <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
        <p className="text-sm text-gray-500 mb-2">
          No ownership records on file for this property yet.
        </p>
        <Link
          to={`/ownership/new?propertyId=${propertyId}`}
          className="inline-block mt-2 text-sm text-teal-700 hover:text-teal-900 font-medium underline"
        >
          Add the first ownership entry →
        </Link>
      </div>
    );
  }

  const totalPercent = directOwnership.reduce((sum, o) => sum + (o.fields.OwnershipPercent ?? 0), 0);

  return (
    <div>
      <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-semibold text-gray-700">Direct owners of {propertyTitle}</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {directOwnership.length} record{directOwnership.length === 1 ? '' : 's'}
              {totalPercent > 0 && ` · ${totalPercent.toFixed(2)}% total`}
            </p>
          </div>
          <Link
            to={`/ownership/new?propertyId=${propertyId}`}
            className="text-xs text-teal-700 hover:text-teal-900 font-medium"
          >
            + Add owner
          </Link>
        </div>
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
            <tr>
              <th className="px-4 py-3 text-left">Entity</th>
              <th className="px-4 py-3 text-left">Type</th>
              <th className="px-4 py-3 text-left">Role</th>
              <th className="px-4 py-3 text-left">Class</th>
              <th className="px-4 py-3 text-right">%</th>
              <th className="px-4 py-3 text-left">Effective</th>
              <th className="px-4 py-3 text-right w-32"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {directOwnership.map((o) => {
              const owner = o.fields.OwnerLookupId ? ownersById.get(String(o.fields.OwnerLookupId)) : null;
              const entityName = owner?.fields.Title ?? o.fields.Title ?? '(unnamed)';
              const isLegacy = !owner;
              const handleUnlink = async (e: React.MouseEvent) => {
                e.stopPropagation();
                const confirmed = window.confirm(
                  `Remove ${entityName} from this property's ownership?\n\nThis deletes the Ownership record only — the Owner entity itself is preserved. ` +
                  `Use this when ownership changes. Action is logged.`
                );
                if (!confirmed) return;
                try {
                  await deleteListItem(LIST_NAMES.Ownership, o.id);
                  await ownership.refetch?.();
                } catch (err) {
                  alert('Failed to unlink: ' + (err instanceof Error ? err.message : String(err)));
                }
              };
              return (
                <tr
                  key={o.id}
                  onClick={() => {
                    if (owner) navigate(`/owners/${owner.id}`);
                    else navigate(`/ownership/${o.id}`);
                  }}
                  className="hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <td className="px-4 py-3 font-medium text-gray-900">
                    {entityName}
                    {isLegacy && (
                      <span className="ml-2 text-[10px] uppercase tracking-wider bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded">
                        legacy
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3">
                    {owner?.fields.OwnerType && (
                      <span className="inline-block px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-100 text-gray-700">
                        {owner.fields.OwnerType}
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3" onClick={(e) => e.stopPropagation()}>
                    <select
                      value={o.fields.RelationshipType ?? ''}
                      onChange={async (e) => {
                        const newValue = e.target.value || null;
                        try {
                          await updateListItem(LIST_NAMES.Ownership, o.id, {
                            RelationshipType: newValue,
                          });
                          await ownership.refetch?.();
                        } catch (err) {
                          alert('Failed to save role: ' + (err instanceof Error ? err.message : String(err)));
                        }
                      }}
                      className={`px-1.5 py-0.5 text-[11px] border rounded cursor-pointer focus:outline-none focus:border-teal-500 hover:border-gray-400 font-semibold ${
                        o.fields.RelationshipType
                          ? RELATIONSHIP_STYLES[o.fields.RelationshipType] || 'bg-white border-gray-200'
                          : 'bg-white border-gray-200'
                      }`}
                      title="Click to change role"
                    >
                      <option value="">—</option>
                      <option value="Managing Member">Managing Member</option>
                      <option value="Sole Member">Sole Member</option>
                      <option value="Member">Member</option>
                      <option value="Owner">Owner</option>
                      <option value="Subsidiary">Subsidiary</option>
                      <option value="Beneficial Owner">Beneficial Owner</option>
                    </select>
                  </td>
                  <td className="px-4 py-3" onClick={(e) => e.stopPropagation()}>
                    <select
                      value={o.fields.MemberClass ?? ''}
                      onChange={async (e) => {
                        const newValue = e.target.value || null;
                        try {
                          await updateListItem(LIST_NAMES.Ownership, o.id, {
                            MemberClass: newValue,
                          });
                          await ownership.refetch?.();
                        } catch (err) {
                          alert('Failed to save class: ' + (err instanceof Error ? err.message : String(err)));
                        }
                      }}
                      className="px-1.5 py-0.5 text-[11px] border border-gray-200 rounded bg-white hover:border-gray-400 focus:outline-none focus:border-teal-500 cursor-pointer"
                      title="Click to set the Member Class for this ownership relationship"
                    >
                      <option value="">—</option>
                      <option value="Class A">Class A</option>
                      <option value="Class B">Class B</option>
                      <option value="Class C">Class C</option>
                      <option value="Class D">Class D</option>
                      <option value="N/A">N/A</option>
                    </select>
                  </td>
                  <td className="px-4 py-3 text-right font-mono-data" onClick={(e) => e.stopPropagation()}>
                    <input
                      type="number"
                      step="0.01"
                      min={0}
                      max={100}
                      value={o.fields.OwnershipPercent ?? ''}
                      onChange={async (e) => {
                        const raw = e.target.value;
                        const newValue = raw === '' ? null : Number(raw);
                        if (raw !== '' && (Number.isNaN(newValue!) || newValue! < 0 || newValue! > 100)) {
                          return;
                        }
                        try {
                          await updateListItem(LIST_NAMES.Ownership, o.id, {
                            OwnershipPercent: newValue,
                          });
                          await ownership.refetch?.();
                        } catch (err) {
                          alert('Failed to save %: ' + (err instanceof Error ? err.message : String(err)));
                        }
                      }}
                      className="w-16 px-1.5 py-0.5 text-[11px] border border-gray-200 rounded text-right hover:border-gray-400 focus:outline-none focus:border-teal-500 bg-white font-mono-data"
                      title="Click to edit ownership percentage"
                    />
                  </td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                    {formatDate(o.fields.EffectiveDate)}
                  </td>
                  <td className="px-4 py-3 text-right">
                    <div className="flex items-center justify-end gap-1">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setEditingOwnershipId(o.id);
                        }}
                        className="text-[11px] text-teal-700 hover:text-teal-900 font-medium px-2 py-1 rounded hover:bg-teal-50 transition-colors"
                        title="Edit all fields on this ownership record"
                      >
                        Edit
                      </button>
                      <button
                        onClick={handleUnlink}
                        className="text-[11px] text-gray-500 hover:text-error font-medium px-2 py-1 rounded hover:bg-red-50 transition-colors"
                        title="Remove this owner from the property"
                      >
                        Unlink
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {totalPercent > 0 && Math.abs(totalPercent - 100) > 0.01 && (
        <div className="mt-3 bg-yellow-50 border border-yellow-200 rounded-lg p-3 flex items-start gap-2">
          <Icon name="alert" size={14} className="text-yellow-700 flex-shrink-0 mt-0.5" />
          <p className="text-xs text-yellow-800">
            Ownership percentages sum to {totalPercent.toFixed(2)}%, not 100%. Either some entries are missing,
            or this property uses a multi-class structure where percentages are within a class (then class totals are what matter).
          </p>
        </div>
      )}

      {editingOwnershipId && (
        <EditOwnershipModal
          ownershipId={editingOwnershipId}
          onClose={() => setEditingOwnershipId(null)}
          onSaved={() => ownership.refetch?.()}
        />
      )}
    </div>
  );
}

// =============================================================================
// Tab: Documents — files linked to this property across libraries
// =============================================================================

/**
 * Document libraries that have a Property lookup column.
 * To add a new library: add it here, ensure it has a "Property" lookup column in SharePoint.
 */
const PROPERTY_LINKED_LIBRARIES = [
  'AMI Certification Letters',
  'DOR Correspondence',
  'DOR Submittal Packages',
  'Land Use Restriction Agreements',
  'Operating Agreements',
  'Org Charts',
  'Property Deeds',
  'Supporting Documentation',
] as const;

interface PropertyDocument {
  id: string;
  itemId: string;
  library: string;
  filename: string;
  webUrl?: string;
  uploadDate?: string;
  size?: number;
  docType?: string;
}

function PropertyDocumentsTab({
  propertyId,
  propertyTitle,
  propertyState,
}: {
  propertyId: string;
  propertyTitle: string;
  propertyState?: 'SC' | 'NC';
}) {
  const [uploadOpen, setUploadOpen] = useState(false);
  const [referenceDocsOpen, setReferenceDocsOpen] = useState(false);

  // Owners + ownership fetch — to identify CAHP entities + property-owner entities for reference sections
  const owners = useSharePointList<Owner>(LIST_NAMES.Owners, { top: 500 });
  const ownership = useSharePointList<Ownership>(LIST_NAMES.Ownership, { top: 500 });

  // Fetch each library in parallel — each useSharePointList is its own hook call
  const lib0 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[0], { top: 500 });
  const lib1 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[1], { top: 500 });
  const lib2 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[2], { top: 500 });
  const lib3 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[3], { top: 500 });
  const lib4 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[4], { top: 500 });
  const lib5 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[5], { top: 500 });
  const lib6 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[6], { top: 500 });
  const lib7 = useSharePointList<DocItemRaw>(PROPERTY_LINKED_LIBRARIES[7], { top: 500 });

  const libraries = [lib0, lib1, lib2, lib3, lib4, lib5, lib6, lib7];
  const loading = libraries.some((l) => l.loading);
  const errors = libraries.filter((l) => l.error).map((l) => l.error!.message);

  const refetchAll = () => libraries.forEach((l) => l.refetch?.());

  const documents = useMemo(() => {
    const docs: PropertyDocument[] = [];
    libraries.forEach((lib, idx) => {
      if (!lib.data) return;
      const libraryName = PROPERTY_LINKED_LIBRARIES[idx];
      lib.data.forEach((item) => {
        if (String(item.fields.PropertyLookupId) !== String(propertyId)) return;
        docs.push({
          id: `${libraryName}:${item.id}`,
          itemId: item.id,
          library: libraryName,
          filename: item.fields.FileLeafRef || item.fields.Title || '(unnamed)',
          webUrl: item.webUrl,
          uploadDate: item.fields.Modified || item.lastModifiedDateTime,
          docType:
            item.fields.LetterType ||
            item.fields.PackageComponent ||
            item.fields.LURAStatus ||
            item.fields.SuppDocType ||
            item.fields.EntityDocType ||
            undefined,
        });
      });
    });
    return docs.sort((a, b) => {
      const da = a.uploadDate ? new Date(a.uploadDate).getTime() : 0;
      const db = b.uploadDate ? new Date(b.uploadDate).getTime() : 0;
      return db - da;
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [propertyId, lib0.data, lib1.data, lib2.data, lib3.data, lib4.data, lib5.data, lib6.data, lib7.data]);

  // Property's actual ownership chain — every owner + parent-owner reachable from this property's Ownership records
  const propertyOwnerChain = useMemo(() => {
    if (!ownership?.data) return new Set<string>();
    const ids = new Set<string>();
    ownership.data.forEach((rel) => {
      if (String(rel.fields.LinkedPropertyLookupId) !== String(propertyId)) return;
      if (rel.fields.OwnerLookupId) ids.add(String(rel.fields.OwnerLookupId));
      if (rel.fields.ParentOwnerLookupId) ids.add(String(rel.fields.ParentOwnerLookupId));
    });
    return ids;
  }, [ownership?.data, propertyId]);

  // CAHP entities to surface on this property — strict chain-based:
  // only CAHP entities that are actually in the property's ownership chain.
  // State filter still applies as a secondary guard (excludes opposite-state entities).
  const cahpEntityIds = useMemo(() => {
    return owners.data
      ?.filter((o) => {
        const title = (o.fields.Title ?? '').toLowerCase();
        const isCahp = title.includes('cahp') || title.includes('carolina affordable housing project');
        if (!isCahp) return false;

        // Chain-based filter — must be in this property's ownership chain
        if (propertyOwnerChain.size > 0 && !propertyOwnerChain.has(String(o.id))) {
          return false;
        }

        // State-scope filter — secondary guard
        if (propertyState && o.fields.OwnerState) {
          if (propertyState === 'SC' && o.fields.OwnerState === 'NC') return false;
          if (propertyState === 'NC' && o.fields.OwnerState === 'SC') return false;
        }
        return true;
      })
      .map((o) => String(o.id)) ?? [];
  }, [owners.data, propertyState, propertyOwnerChain]);

  // Identify property's direct-owner entity IDs (the LLCs that hold this property)
  const propertyOwnerIds = useMemo(() => {
    if (!ownership?.data) return [] as string[];
    const ids = new Set<string>();
    ownership.data.forEach((rel) => {
      if (String(rel.fields.LinkedPropertyLookupId) !== String(propertyId)) return;
      if (rel.fields.OwnerLookupId) {
        const ownerId = String(rel.fields.OwnerLookupId);
        // Don't double-count CAHP entities — they get their own section
        if (!cahpEntityIds.includes(ownerId)) ids.add(ownerId);
      }
    });
    return Array.from(ids);
  }, [ownership?.data, propertyId, cahpEntityIds]);

  if (loading) return <TabLoading label={`documents across ${PROPERTY_LINKED_LIBRARIES.length} libraries`} />;
  if (errors.length > 0) return <TabError error={new Error(errors[0])} />;

  const handleUploadSuccess = () => {
    setUploadOpen(false);
    refetchAll();
  };

  return (
    <div>
      {/* Property documents — primary content */}
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-gray-700">
          {documents.length === 0
            ? 'No documents tagged to this property yet'
            : `${documents.length} document${documents.length === 1 ? '' : 's'} tagged to this property`}
        </h3>
        <button
          onClick={() => setUploadOpen(true)}
          className="text-xs text-teal-700 hover:text-teal-900 font-medium flex items-center gap-1"
        >
          <Icon name="plus" size={12} />
          Upload Document
        </button>
      </div>

      {documents.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500 mb-3">No documents tagged to this property yet.</p>
          <button
            onClick={() => setUploadOpen(true)}
            className="bg-teal-700 hover:bg-teal-900 text-white px-3 py-1.5 rounded-md text-sm font-medium inline-flex items-center gap-1.5"
          >
            <Icon name="plus" size={14} />
            Upload First Document
          </button>
          <p className="text-xs text-gray-400 mt-3">
            Pick a library + a file; PropertyID metadata is set automatically.
          </p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <div className="px-4 py-2 border-b border-gray-100 bg-gray-50">
            <p className="text-xs text-gray-500">
              Across {PROPERTY_LINKED_LIBRARIES.length} libraries · click filename to open in SharePoint
            </p>
          </div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-600 uppercase tracking-wider">
              <tr>
                <th className="px-4 py-3 text-left">Filename</th>
                <th className="px-4 py-3 text-left">Library</th>
                <th className="px-4 py-3 text-left">Type</th>
                <th className="px-4 py-3 text-left">Modified</th>
                <th className="px-4 py-3 text-right w-16"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {documents.map((doc) => {
                const handleDelete = async (e: React.MouseEvent) => {
                  e.preventDefault();
                  e.stopPropagation();
                  const ok = window.confirm(
                    `Permanently delete "${doc.filename}"?\n\nThis removes the file from SharePoint and cannot be undone. Action is logged.`
                  );
                  if (!ok) return;
                  try {
                    await deleteListItem(doc.library as typeof PROPERTY_LINKED_LIBRARIES[number], doc.itemId);
                    // Refetch libraries — find the matching one
                    const idx = PROPERTY_LINKED_LIBRARIES.indexOf(doc.library as typeof PROPERTY_LINKED_LIBRARIES[number]);
                    if (idx >= 0) await libraries[idx]?.refetch?.();
                  } catch (err) {
                    alert('Failed to delete: ' + (err instanceof Error ? err.message : String(err)));
                  }
                };
                return (
                <tr key={doc.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3 font-medium text-gray-900">
                    {doc.webUrl ? (
                      <a
                        href={doc.webUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-teal-700 hover:text-teal-900 underline"
                      >
                        {doc.filename}
                      </a>
                    ) : (
                      doc.filename
                    )}
                  </td>
                  <td className="px-4 py-3 text-gray-700 text-xs">{doc.library}</td>
                  <td className="px-4 py-3 text-gray-700 text-xs">{doc.docType || '—'}</td>
                  <td className="px-4 py-3 text-gray-700 font-mono-data text-xs">
                    {formatDate(doc.uploadDate)}
                  </td>
                  <td className="px-4 py-3 text-right">
                    <button
                      onClick={handleDelete}
                      className="text-[11px] text-gray-500 hover:text-error font-medium px-2 py-1 rounded hover:bg-red-50 transition-colors"
                      title="Delete this document permanently"
                    >
                      Delete
                    </button>
                  </td>
                </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* Reference documents — collapsible, shown below property docs */}
      {(cahpEntityIds.length > 0 || propertyOwnerIds.length > 0) && (
        <div className="mt-6 bg-white border border-gray-200 rounded-lg shadow-card overflow-hidden">
          <button
            onClick={() => setReferenceDocsOpen((v) => !v)}
            className="w-full px-4 py-3 flex items-center justify-between hover:bg-gray-50 transition-colors text-left"
          >
            <div className="flex items-center gap-2">
              <Icon
                name={referenceDocsOpen ? 'check' : 'plus'}
                size={12}
                className="text-gray-400 flex-shrink-0"
              />
              <h3 className="text-sm font-semibold text-teal-900">Reference Documents</h3>
              <span className="text-[11px] text-gray-500">
                Entity-level docs from this property's ownership chain (CAHP entities + owner LLC formation docs).
              </span>
            </div>
            <span className="text-[11px] text-teal-700 font-medium">
              {referenceDocsOpen ? 'Hide' : 'Show'}
            </span>
          </button>

          {referenceDocsOpen && (
            <div className="border-t border-gray-100">
              {/* CAHP entity reference docs — only entities in this property's ownership chain */}
              {cahpEntityIds.length > 0 && (
                <EntityDocumentsSection
                  ownerIds={cahpEntityIds}
                  title="CAHP Entity Documents"
                  subtitle="Filtered to entities in this property's ownership chain only."
                  variant="inline"
                  useCahpEntityLibrary
                />
              )}

              {/* Property-owner reference docs */}
              {propertyOwnerIds.length > 0 && (
                <EntityDocumentsSection
                  ownerIds={propertyOwnerIds}
                  title="Property-Owner Entity Documents"
                  subtitle="Formation docs for the LLC that holds this property — EIN, Articles, COE, Cert of Authorization."
                  variant="inline"
                />
              )}

              {cahpEntityIds.length === 0 && propertyOwnerIds.length === 0 && (
                <div className="px-4 py-6 text-center text-xs text-gray-500">
                  No reference documents found. Make sure this property has Ownership records linking
                  it to its owner entities.
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {uploadOpen && (
        <UploadDocumentModal
          propertyId={propertyId}
          propertyTitle={propertyTitle}
          onSuccess={handleUploadSuccess}
          onClose={() => setUploadOpen(false)}
        />
      )}
    </div>
  );
}

// Raw shape returned by Graph for a document library item
interface DocItemRaw {
  id: string;
  webUrl?: string;
  lastModifiedDateTime: string;
  fields: {
    Title?: string;
    FileLeafRef?: string;
    PropertyLookupId?: string;
    Modified?: string;
    LetterType?: string;
    PackageComponent?: string;
    LURAStatus?: string;
    SuppDocType?: string;
    EntityDocType?: string;
  };
}

// =============================================================================
// Tab: Notes — per-property notes thread (PR-08d)
// =============================================================================

function PropertyNotesTab({ propertyId, propertyTitle }: { propertyId: string; propertyTitle: string }) {
  const { data: allNotes, loading, error, refetch } = useSharePointList<PropertyNote>(
    LIST_NAMES.PropertyNotes,
    { top: 500 }
  );

  const [newNote, setNewNote] = useState('');
  const [posting, setPosting] = useState(false);
  const [postError, setPostError] = useState<string | null>(null);

  const propertyNotes = useMemo(() => {
    if (!allNotes) return [];
    return allNotes
      .filter((n) => String(n.fields.PropertyLookupId) === String(propertyId))
      .sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime());
  }, [allNotes, propertyId]);

  const handleAdd = async () => {
    const trimmed = newNote.trim();
    if (!trimmed) return;
    setPosting(true);
    setPostError(null);
    try {
      // Title gets first 80 chars of body since SP requires Title to be non-empty.
      // Frontend never displays Title — only NoteBody, author, timestamp.
      const title = trimmed.length > 80 ? trimmed.slice(0, 77) + '…' : trimmed;
      await createListItem(LIST_NAMES.PropertyNotes, {
        Title: title,
        NoteBody: trimmed,
        PropertyLookupId: propertyId,
      });
      setNewNote('');
      refetch();
    } catch (err) {
      setPostError(err instanceof Error ? err.message : String(err));
    } finally {
      setPosting(false);
    }
  };

  if (loading) return <TabLoading label="notes" />;
  if (error) {
    // Most likely cause: Property Notes list doesn't exist yet — provisioning script wasn't run.
    return (
      <div className="bg-red-50 border border-red-200 rounded-lg p-6">
        <div className="font-semibold text-error mb-2 flex items-center gap-2">
          <Icon name="alert" size={18} />
          Notes list unavailable
        </div>
        <p className="text-sm text-red-700 mb-2">
          Most likely cause: the Property Notes SharePoint list hasn't been provisioned yet.
        </p>
        <p className="text-xs text-red-600 font-mono-data">{error.message}</p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* New note input */}
      <div className="bg-white border border-gray-200 rounded-lg shadow-card p-4">
        <h3 className="text-sm font-semibold text-gray-700 mb-2">Add a note</h3>
        <textarea
          value={newNote}
          onChange={(e) => setNewNote(e.target.value)}
          rows={3}
          placeholder={`Note about ${propertyTitle}…`}
          disabled={posting}
          className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 resize-y disabled:opacity-50"
        />
        {postError && (
          <div className="mt-2 bg-red-50 border border-red-200 rounded p-2 text-sm text-error font-mono-data text-xs">
            {postError}
          </div>
        )}
        <div className="mt-2 flex items-center justify-between gap-3">
          <p className="text-xs text-gray-400">
            Notes are timestamped and attributed automatically. Logged to the Audit Log.
          </p>
          <button
            onClick={handleAdd}
            disabled={!newNote.trim() || posting}
            className="px-3 py-1.5 bg-teal-700 hover:bg-teal-900 text-white rounded-md text-sm font-medium disabled:opacity-50 disabled:cursor-not-allowed transition-colors flex-shrink-0 flex items-center gap-1.5"
          >
            {posting && <div className="w-3 h-3 rounded-full border-2 border-white border-r-transparent animate-spin" />}
            {posting ? 'Posting…' : 'Post Note'}
          </button>
        </div>
      </div>

      {/* Notes list */}
      {propertyNotes.length === 0 ? (
        <div className="bg-white border border-gray-200 rounded-lg p-8 text-center shadow-card">
          <p className="text-sm text-gray-500">No notes yet for this property. Add the first one above.</p>
        </div>
      ) : (
        <div className="bg-white border border-gray-200 rounded-lg shadow-card divide-y divide-gray-100">
          <div className="px-4 py-2 bg-gray-50 border-b border-gray-100">
            <h3 className="text-xs font-semibold text-gray-600 uppercase tracking-wider">
              {propertyNotes.length} note{propertyNotes.length === 1 ? '' : 's'} · most recent first
            </h3>
          </div>
          {propertyNotes.map((note) => (
            <div key={note.id} className="p-4">
              <div className="flex items-center gap-2 mb-1.5 text-xs">
                <span className="font-semibold text-gray-700">
                  {note.createdBy?.user?.displayName ?? 'Unknown author'}
                </span>
                <span className="text-gray-400">·</span>
                <span className="font-mono-data text-gray-500">
                  {formatNoteTimestamp(note.createdDateTime)}
                </span>
              </div>
              <p className="text-sm text-gray-800 whitespace-pre-wrap">
                {note.fields.NoteBody || note.fields.Title || '(empty note)'}
              </p>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function formatNoteTimestamp(iso: string): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const now = Date.now();
  const diffMs = now - d.getTime();
  const diffMin = Math.floor(diffMs / 60000);
  const diffHr = Math.floor(diffMin / 60);
  const diffDay = Math.floor(diffHr / 24);

  if (diffMin < 1) return 'just now';
  if (diffMin < 60) return `${diffMin}m ago`;
  if (diffHr < 24) return `${diffHr}h ago`;
  if (diffDay < 7) return `${diffDay}d ago`;
  return d.toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
}

// =============================================================================
// Shared tab states
// =============================================================================

function TabLoading({ label }: { label: string }) {
  return (
    <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
      <div className="inline-flex items-center gap-3 text-gray-500">
        <div className="w-4 h-4 rounded-full border-2 border-teal-500 border-r-transparent animate-spin" />
        <span className="text-sm">Loading {label}…</span>
      </div>
    </div>
  );
}

function TabError({ error }: { error: Error }) {
  return (
    <div className="bg-red-50 border border-red-200 rounded-lg p-6">
      <div className="font-semibold text-error mb-2 flex items-center gap-2">
        <Icon name="alert" size={18} />
        Failed to load
      </div>
      <p className="text-sm text-red-700 font-mono-data text-xs">{error.message}</p>
    </div>
  );
}

function BreadcrumbBar({ propertyName }: { propertyName?: string }) {
  return (
    <nav className="mb-4 text-sm">
      <Link to="/properties" className="text-teal-700 hover:text-teal-900 font-medium">
        ← Properties
      </Link>
      {propertyName && (
        <>
          <span className="mx-2 text-gray-400">/</span>
          <span className="text-gray-700">{propertyName}</span>
        </>
      )}
    </nav>
  );
}

function Section({
  title, children, fullWidth,
}: {
  title: string;
  children: React.ReactNode;
  fullWidth?: boolean;
}) {
  return (
    <div className={`bg-white border border-gray-200 rounded-lg p-5 shadow-card ${fullWidth ? 'lg:col-span-2' : ''}`}>
      <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">{title}</h3>
      <div className="space-y-2.5">{children}</div>
    </div>
  );
}

type FieldValue = string | number | null | undefined;
type FieldType = 'text' | 'textarea' | 'number' | 'date' | 'choice';

function EditableField({
  label, value, editing, onChange,
  type = 'text', choices, mono, required, rows = 3, hideLabel,
}: {
  label: string;
  value: FieldValue;
  editing: boolean;
  onChange: (v: FieldValue) => void;
  type?: FieldType;
  choices?: readonly string[];
  mono?: boolean;
  required?: boolean;
  rows?: number;
  hideLabel?: boolean;
}) {
  if (!editing) {
    let displayValue: React.ReactNode = value;
    if (type === 'date' && value) displayValue = formatDate(value as string);
    if (type === 'textarea' && value) {
      displayValue = <span className="whitespace-pre-wrap leading-relaxed block">{String(value)}</span>;
    }
    return (
      <div className="flex items-start gap-3">
        {!hideLabel && <dt className="text-sm text-gray-500 w-44 flex-shrink-0">{label}</dt>}
        <dd className={`text-sm text-gray-900 flex-1 ${mono ? 'font-mono-data' : ''}`}>
          {value === null || value === undefined || value === '' ? (
            <span className="text-gray-300">—</span>
          ) : displayValue}
        </dd>
      </div>
    );
  }

  const inputClass = 'w-full px-2 py-1 border border-gray-300 rounded text-sm focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500';

  let input: React.ReactNode;
  if (type === 'textarea') {
    input = (
      <textarea value={String(value ?? '')} onChange={(e) => onChange(e.target.value)} rows={rows} className={`${inputClass} resize-y`} />
    );
  } else if (type === 'number') {
    input = (
      <input type="number" value={value == null || value === '' ? '' : value}
        onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}
        className={`${inputClass} ${mono ? 'font-mono-data' : ''}`} />
    );
  } else if (type === 'date') {
    const dateValue = value ? new Date(value as string).toISOString().slice(0, 10) : '';
    input = (
      <input type="date" value={dateValue}
        onChange={(e) => onChange(e.target.value ? new Date(e.target.value).toISOString() : null)}
        className={`${inputClass} font-mono-data`} />
    );
  } else if (type === 'choice') {
    input = (
      <select value={String(value ?? '')} onChange={(e) => onChange(e.target.value || null)}
        className={`${inputClass} bg-white`}>
        <option value="">— none —</option>
        {choices?.map((c) => (<option key={c} value={c}>{c}</option>))}
      </select>
    );
  } else {
    input = (
      <input type="text" value={String(value ?? '')} onChange={(e) => onChange(e.target.value)}
        className={`${inputClass} ${mono ? 'font-mono-data' : ''}`} required={required} />
    );
  }

  return (
    <div className="flex items-start gap-3">
      {!hideLabel && (
        <label className="text-sm text-gray-500 w-44 flex-shrink-0 pt-1">
          {label}{required && <span className="text-error ml-0.5">*</span>}
        </label>
      )}
      <div className="flex-1">{input}</div>
    </div>
  );
}

function formatDate(iso: string | undefined | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
