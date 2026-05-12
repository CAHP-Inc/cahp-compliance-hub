import { Routes, Route, Navigate } from 'react-router-dom';
import { AppShell } from './components/layout/AppShell';
import { SessionProvider } from './lib/session';
import { SignInGate } from './components/auth/SignInGate';
import { MyDay } from './pages/MyDay';
import { Properties } from './pages/Properties';
import { PropertyDetail } from './pages/PropertyDetail';
import { PlaceholderPage } from './pages/PlaceholderPage';

function App() {
  return (
    <SessionProvider>
      <SignInGate>
        <AppShell>
          <Routes>
            <Route path="/" element={<MyDay />} />
            <Route path="/outstanding" element={<PlaceholderPage title="Outstanding Items" icon="alert" plannedPR="PR-07" />} />
            <Route path="/portfolio" element={<PlaceholderPage title="Portfolio" icon="grid" plannedPR="PR-05" />} />
            <Route path="/properties" element={<Properties />} />
            <Route path="/properties/:id" element={<PropertyDetail />} />
            <Route path="/owners" element={<PlaceholderPage title="Owners" icon="star" plannedPR="PR-06" />} />
            <Route path="/cahp-entity" element={<PlaceholderPage title="CAHP Entity" icon="star" plannedPR="PR-06" />} />
            <Route path="/submittals" element={<PlaceholderPage title="Submittals" icon="file" plannedPR="Phase 2" />} />
            <Route path="/correspondence" element={<PlaceholderPage title="DOR Correspondence" icon="mail" plannedPR="Phase 2" />} />
            <Route path="/comms" element={<PlaceholderPage title="Owner Communications" icon="mail" plannedPR="Phase 3" />} />
            <Route path="/compliance" element={<PlaceholderPage title="Compliance" icon="calendar" plannedPR="Phase 3" />} />
            <Route path="/billing" element={<PlaceholderPage title="Billing & Disbursements" icon="dollar" plannedPR="Phase 3" />} />
            <Route path="/documents" element={<PlaceholderPage title="Documents" icon="folder" plannedPR="Phase 2" />} />
            <Route path="/untagged" element={<PlaceholderPage title="Untagged Documents" icon="alert" plannedPR="Phase 2" />} />
            <Route path="/reports" element={<PlaceholderPage title="Reports" icon="file" plannedPR="Phase 3" />} />
            <Route path="/audit" element={<PlaceholderPage title="Audit Log" icon="history" plannedPR="PR-07" />} />
            <Route path="/settings" element={<PlaceholderPage title="Settings" icon="settings" plannedPR="Phase 3" />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </AppShell>
      </SignInGate>
    </SessionProvider>
  );
}

export default App;
