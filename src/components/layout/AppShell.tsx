import { useState, ReactNode } from 'react';
import { Header } from './Header';
import { SidebarNav } from './SidebarNav';
import { FilingFreezeBanner, useFilingFreezeStatus } from './FilingFreezeBanner';

interface AppShellProps {
  children: ReactNode;
}

export function AppShell({ children }: AppShellProps) {
  const [navOpen, setNavOpen] = useState(false);
  const freezeStatus = useFilingFreezeStatus();
  const mainPad = freezeStatus ? 'pt-[6rem]' : 'pt-14';

  return (
    <div className="min-h-screen bg-gray-50">
      <Header onOpenNav={() => setNavOpen(true)} />
      {freezeStatus && <FilingFreezeBanner status={freezeStatus} />}
      <SidebarNav open={navOpen} onClose={() => setNavOpen(false)} />
      <main className={mainPad}>
        <div className="max-w-[1400px] mx-auto px-6 py-6">
          {children}
        </div>
      </main>
    </div>
  );
}
