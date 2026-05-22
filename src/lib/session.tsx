import { createContext, useContext, useEffect, useMemo, useState, ReactNode } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import type { Role } from './permissions';
import { lookupRole, loadAccessList, getUserInitials, extractOrgFromEmail } from './roleMap';

export interface User {
  id: string;
  name: string;
  email: string;
  initials: string;
  org: string;
}

interface SessionState {
  user: User | null;
  /** Effective role for the current view (may be a dev override) */
  role: Role | null;
  /** Real role from the M365 → role map; null if not on access list */
  realRole: Role | null;
  isAuthenticated: boolean;
  /** Authenticated AND on the role map */
  isAuthorized: boolean;
  /** True while we're still loading the Access List on first sign-in.
   *  SignInGate uses this to show a "Checking access…" splash instead of
   *  flashing "Access denied" before the fetch lands. */
  accessLoading: boolean;
  /** Force a re-fetch of the Access List (used after Settings → Access List saves). */
  refreshAccess: () => Promise<void>;
  /** Dev-only: change the view-as role. Undefined in production builds. */
  setDevRoleOverride?: (role: Role | null) => void;
}

const SessionContext = createContext<SessionState | null>(null);

export function SessionProvider({ children }: { children: ReactNode }) {
  const { accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [devOverride, setDevOverride] = useState<Role | null>(null);
  // Bump on access-list reload to recompute role
  const [accessVersion, setAccessVersion] = useState(0);
  const [accessLoading, setAccessLoading] = useState(false);

  // Load the Access List from SharePoint as soon as the user authenticates.
  // Before this resolves, lookupRole() falls back to the hardcoded map — so
  // the original team can still get in even if SP is unreachable.
  useEffect(() => {
    if (!isAuthenticated || accounts.length === 0) return;
    let cancelled = false;
    setAccessLoading(true);
    loadAccessList()
      .catch(() => {
        // loadAccessList swallows its own errors and sets the fallback;
        // catch here is just belt-and-suspenders.
      })
      .finally(() => {
        if (cancelled) return;
        setAccessLoading(false);
        setAccessVersion((v) => v + 1);
      });
    return () => { cancelled = true; };
  }, [isAuthenticated, accounts]);

  const refreshAccess = async () => {
    setAccessLoading(true);
    try {
      await loadAccessList();
    } finally {
      setAccessLoading(false);
      setAccessVersion((v) => v + 1);
    }
  };

  const state = useMemo<SessionState>(() => {
    if (!isAuthenticated || accounts.length === 0) {
      return {
        user: null,
        role: null,
        realRole: null,
        isAuthenticated: false,
        isAuthorized: false,
        accessLoading: false,
        refreshAccess,
        setDevRoleOverride: import.meta.env.DEV ? setDevOverride : undefined,
      };
    }

    const account = accounts[0];
    // For Azure AD accounts, `username` is the UPN / email
    const email = account.username;
    const name = account.name || email;
    const realRole = lookupRole(email);
    const effectiveRole = import.meta.env.DEV && devOverride ? devOverride : realRole;

    const user: User = {
      id: account.localAccountId,
      name,
      email,
      initials: getUserInitials(name, email),
      org: extractOrgFromEmail(email),
    };

    return {
      user,
      role: effectiveRole,
      realRole,
      isAuthenticated: true,
      isAuthorized: realRole !== null,
      accessLoading,
      refreshAccess,
      setDevRoleOverride: import.meta.env.DEV ? setDevOverride : undefined,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAuthenticated, accounts, devOverride, accessVersion, accessLoading]);

  return <SessionContext.Provider value={state}>{children}</SessionContext.Provider>;
}

export function useSession(): SessionState {
  const ctx = useContext(SessionContext);
  if (!ctx) throw new Error('useSession must be used inside SessionProvider');
  return ctx;
}
