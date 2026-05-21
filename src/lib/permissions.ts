/**
 * Role-based access control.
 *
 * Three roles per spec Section 13:
 *   - Admin: Bryan, Stan, Brandy. Full access.
 *   - Contributor: Lori, Cara. Edit operational data, no Settings, no Billing detail.
 *   - Accounting: Chris. Full Billing, view-only on Properties/Submittals, no Owners/CAHP/Correspondence.
 *
 * Permissions are enforced both in the UI (hiding nav items, disabling actions) AND
 * at the data layer (SharePoint List permissions). This module is the UI-side check.
 */

export type Role = 'Admin' | 'Contributor' | 'Accounting';

export type ModuleId =
  | 'myday'
  | 'portfolio'
  | 'properties'
  | 'owners'
  | 'cahp-entity'
  | 'contacts'
  | 'submittals'
  | 'correspondence'
  | 'comms'
  | 'outstanding'
  | 'compliance'
  | 'billing'
  | 'documents'
  | 'untagged'
  | 'reports'
  | 'audit'
  | 'settings'
  | 'property-detail'
  | 'owner-detail';

export type Action = 'view' | 'create' | 'edit' | 'delete' | 'approve' | 'upload' | 'export';

interface RolePermission {
  label: string;
  description: string;
  color: string; // Tailwind class for badge
  views: ModuleId[];
  actions: Partial<Record<ModuleId, Action[]>>;
}

export const ROLE_PERMISSIONS: Record<Role, RolePermission> = {
  Admin: {
    label: 'Admin / Owner',
    description: 'Full access to every module and action.',
    color: 'bg-gold-500 text-teal-900',
    views: [
      'myday', 'portfolio', 'properties', 'owners', 'cahp-entity', 'contacts',
      'submittals', 'correspondence', 'comms', 'outstanding',
      'compliance', 'billing', 'documents', 'untagged',
      'reports', 'audit', 'settings',
      'property-detail', 'owner-detail',
    ],
    actions: {}, // Admin has all actions implicitly; checked by canDo's fallback
  },
  Contributor: {
    label: 'Contributor',
    description: 'Edit operational data. No Settings, no Billing detail, no destructive actions.',
    color: 'bg-teal-700 text-white',
    views: [
      'myday', 'portfolio', 'properties', 'owners', 'cahp-entity', 'contacts',
      'submittals', 'correspondence', 'comms', 'outstanding',
      'compliance', 'documents', 'untagged',
      'reports', 'audit',
      'property-detail', 'owner-detail',
    ],
    actions: {
      properties: ['view', 'create', 'edit', 'upload'],
      owners: ['view', 'create', 'edit'],
      'cahp-entity': ['view', 'edit'],
      contacts: ['view', 'create', 'edit'],
      submittals: ['view', 'create', 'edit'],
      correspondence: ['view', 'create', 'edit'],
      comms: ['view', 'create', 'edit'],
      outstanding: ['view', 'create', 'edit'],
      compliance: ['view', 'edit'],
      documents: ['view', 'upload'],
      untagged: ['view', 'edit'],
      reports: ['view', 'export'],
      audit: ['view'],
    },
  },
  Accounting: {
    label: 'Accounting',
    description: 'Full Billing and Reports. View-only on Properties and Submittals. No Owners, no Correspondence.',
    color: 'bg-info text-white',
    views: [
      'myday', 'portfolio', 'properties',
      'submittals', 'outstanding',
      'billing', 'documents',
      'reports', 'comms', 'audit',
      'property-detail',
    ],
    actions: {
      portfolio: ['view'],
      properties: ['view'],
      submittals: ['view'],
      outstanding: ['view', 'edit'], // own assignments only — enforced at data layer
      billing: ['view', 'create', 'edit', 'export'],
      documents: ['view'],
      comms: ['view'],
      reports: ['view', 'export'],
      audit: ['view'],
    },
  },
};

/**
 * Can this role view this module? Used to filter nav items and gate routes.
 */
export function canView(role: Role, moduleId: ModuleId): boolean {
  return ROLE_PERMISSIONS[role]?.views.includes(moduleId) ?? false;
}

/**
 * Can this role perform this action on this module?
 * Admin has all actions implicitly. Other roles must have the action explicitly granted.
 */
export function canDo(role: Role, moduleId: ModuleId, action: Action): boolean {
  if (role === 'Admin') return true;
  const allowed = ROLE_PERMISSIONS[role]?.actions[moduleId];
  return allowed?.includes(action) ?? false;
}
