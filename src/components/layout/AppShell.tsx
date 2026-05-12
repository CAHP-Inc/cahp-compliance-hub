import { useState, ReactNode } from 'react';
import { Header } from './Header';
import { SidebarNav } from './SidebarNav';

interface AppShellProps {
  children: ReactNode;
}

export function AppShell({ children }: AppShellProps) {
  const [navOpen, setNavOpen] = useState(false);

  return (
    <div className="min-h-screen bg-gray-50">
      <Header onOpenNav={() => setNavOpen(true)} />
      <SidebarNav open={navOpen} onClose={() => setNavOpen(false)} />
      <main className="pt-14">
        <div className="max-w-[1400px] mx-auto px-6 py-6">
          {children}
        </div>
      </main>
    </div>
  );
}
