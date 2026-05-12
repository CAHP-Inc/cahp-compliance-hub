/**
 * Session — currently a stub for PR-01.
 *
 * In PR-02, this becomes a real MSAL-backed hook reading the signed-in M365 user
 * and looking up their role from the Users SharePoint List.
 *
 * For now, exposes a static "Brandy Turner / Admin" session plus a setter for testing
 * role-based UI in development.
 */

import { createContext, useContext, useState, ReactNode } from 'react';
import type { Role } from './permissions';

export interface User {
  id: string;
  name: string;
  email: string;
  initials: string;
  org: string;
}

interface SessionState {
  user: User;
  role: Role;
  setRole: (role: Role) => void; // dev-only — remove in PR-02
}

const DEFAULT_USER: User = {
  id: 'stub-brandy',
  name: 'Brandy Turner',
  email: 'brandy@newshire.com',
  initials: 'BT',
  org: 'NewShire',
};

const SessionContext = createContext<SessionState | null>(null);

export function SessionProvider({ children }: { children: ReactNode }) {
  const [role, setRole] = useState<Role>('Admin');
  return (
    <SessionContext.Provider value={{ user: DEFAULT_USER, role, setRole }}>
      {children}
    </SessionContext.Provider>
  );
}

export function useSession(): SessionState {
  const ctx = useContext(SessionContext);
  if (!ctx) throw new Error('useSession must be used inside SessionProvider');
  return ctx;
}
