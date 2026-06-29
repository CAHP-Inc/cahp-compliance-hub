import { Routes, Route, Navigate } from 'react-router-dom';
import { AppShell } from './components/layout/AppShell';
import { SessionProvider } from './lib/session';
import { SignInGate } from './components/auth/SignInGate';
import { MyDay } from './pages/MyDay';
import { Properties } from './pages/Properties';
import { PropertyDetail } from './pages/PropertyDetail';
import { PropertyNew } from './pages/PropertyNew';
import { Portfolio } from './pages/Portfolio';
import { Owners } from './pages/Owners';
import { OwnerDetail } from './pages/OwnerDetail';
import { OwnerNew } from './pages/OwnerNew';
import { CAHPEntity } from './pages/CAHPEntity';
import { Contacts } from './pages/Contacts';
import { Submittals } from './pages/Submittals';
import { SubmittalDetail } from './pages/SubmittalDetail';
import { CorrespondencePage } from './pages/CorrespondencePage';
import { CorrespondenceDetail } from './pages/CorrespondenceDetail';
import { OutstandingItems } from './pages/OutstandingItems';
import { OutstandingItemDetail } from './pages/OutstandingItemDetail';
import { UntaggedDocuments } from './pages/UntaggedDocuments';
import { OwnerCommunicationsPage } from './pages/OwnerCommunicationsPage';
import { OwnerCommunicationDetail } from './pages/OwnerCommunicationDetail';
import { ReportsPage } from './pages/ReportsPage';
import { NotificationsPage } from './pages/NotificationsPage';
import { DocumentsPage } from './pages/DocumentsPage';
import { SettingsPage } from './pages/SettingsPage';
import { Compliance } from './pages/Compliance';
import { ComplianceDeadlineDetail } from './pages/ComplianceDeadlineDetail';
import { Ownership } from './pages/Ownership';
import { OwnershipDetail } from './pages/OwnershipDetail';
import { OwnershipNew } from './pages/OwnershipNew';
import { Audit } from './pages/Audit';
import { BillingPage } from './pages/BillingPage';
import { BillingDetail } from './pages/BillingDetail';
import { DisbursementDetail } from './pages/DisbursementDetail';

function App() {
  return (
    <SessionProvider>
      <SignInGate>
        <AppShell>
          <Routes>
            <Route path="/" element={<MyDay />} />
            <Route path="/outstanding" element={<OutstandingItems />} />
            <Route path="/outstanding-items" element={<OutstandingItems />} />
            <Route path="/outstanding-items/:id" element={<OutstandingItemDetail />} />
            <Route path="/untagged-documents" element={<UntaggedDocuments />} />
            <Route path="/portfolio" element={<Portfolio />} />
            <Route path="/properties" element={<Properties />} />
            <Route path="/properties/new" element={<PropertyNew />} />
            <Route path="/properties/:id" element={<PropertyDetail />} />
            <Route path="/owners" element={<Owners />} />
            <Route path="/owners/new" element={<OwnerNew />} />
            <Route path="/owners/:id" element={<OwnerDetail />} />
            <Route path="/cahp-entity" element={<CAHPEntity />} />
            <Route path="/contacts" element={<Contacts />} />
            <Route path="/ownership" element={<Ownership />} />
            <Route path="/ownership/new" element={<OwnershipNew />} />
            <Route path="/ownership/:id" element={<OwnershipDetail />} />
            <Route path="/submittals" element={<Submittals />} />
            <Route path="/submittals/:id" element={<SubmittalDetail />} />
            <Route path="/correspondence" element={<CorrespondencePage />} />
            <Route path="/correspondence/:id" element={<CorrespondenceDetail />} />
            <Route path="/comms" element={<OwnerCommunicationsPage />} />
            <Route path="/comms/:id" element={<OwnerCommunicationDetail />} />
            <Route path="/compliance" element={<Compliance />} />
            <Route path="/compliance/:id" element={<ComplianceDeadlineDetail />} />
            <Route path="/billing" element={<BillingPage />} />
            <Route path="/billing/invoices/:id" element={<BillingDetail />} />
            <Route path="/billing/disbursements/:id" element={<DisbursementDetail />} />
            <Route path="/documents" element={<DocumentsPage />} />
            <Route path="/untagged" element={<UntaggedDocuments />} />
            <Route path="/reports" element={<ReportsPage />} />
            <Route path="/notifications" element={<NotificationsPage />} />
            <Route path="/audit" element={<Audit />} />
            <Route path="/settings" element={<SettingsPage />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </AppShell>
      </SignInGate>
    </SessionProvider>
  );
}

export default App;
