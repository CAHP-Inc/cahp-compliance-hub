import React from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter } from 'react-router-dom';
import { MsalProvider } from '@azure/msal-react';
import App from './App';
import { msalInstance } from './lib/auth/msalConfig';
import './styles/globals.css';

const basePath = import.meta.env.VITE_BASE_PATH || '/';

/**
 * Bootstrap sequence:
 *   1. Initialize MSAL (required before any other MSAL API calls)
 *   2. Handle redirect promise — resolves with account info after returning from sign-in
 *   3. If no redirect response, pick up any existing cached account as the active one
 *   4. Mount React tree wrapped in MsalProvider
 */
async function bootstrap() {
  await msalInstance.initialize();

  try {
    const response = await msalInstance.handleRedirectPromise();
    if (response?.account) {
      msalInstance.setActiveAccount(response.account);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
      }
    }
  } catch (err) {
    console.error('[MSAL] Redirect handling failed:', err);
  }

  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <BrowserRouter basename={basePath}>
          <App />
        </BrowserRouter>
      </MsalProvider>
    </React.StrictMode>
  );
}

bootstrap();
