import { createContext, useContext, useMemo, useState, ReactNode } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import type { Role } from './permissions';
import { lookupRole, getUserInitials, extractOrgFromEmail } from './roleMap';

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
  /** Dev-only: change the view-as role. Undefined in production builds. */
  setDevRoleOverride?: (role: Role | null) => void;
}

const SessionContext = createContext<SessionState | null>(null);

export function SessionProvider({ children }: { children: ReactNode }) {
  const { accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [devOverride, setDevOverride] = useState<Role | null>(null);

  const state = useMemo<SessionState>(() => {
    if (!isAuthenticated || accounts.length === 0) {
      return {
        user: null,
        role: null,
        realRole: null,
        isAuthenticated: false,
        isAuthorized: false,
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
      setDevRoleOverride: import.meta.env.DEV ? setDevOverride : undefined,
    };
  }, [isAuthenticated, accounts, devOverride]);

  return <SessionContext.Provider value={state}>{children}</SessionContext.Provider>;
}

export function useSession(): SessionState {
  const ctx = useContext(SessionContext);
  if (!ctx) throw new Error('useSession must be used inside SessionProvider');
  return ctx;
}
