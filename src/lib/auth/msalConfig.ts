import { Configuration, LogLevel, PublicClientApplication } from '@azure/msal-browser';

const clientId = import.meta.env.VITE_AZURE_CLIENT_ID;
const tenantId = import.meta.env.VITE_AZURE_TENANT_ID;

if (!clientId || !tenantId) {
  console.error(
    '[MSAL] Missing required env vars: VITE_AZURE_CLIENT_ID and VITE_AZURE_TENANT_ID. ' +
    'Locally: copy .env.example to .env.local and fill in values from your Azure AD app registration. ' +
    'Production: set these as GitHub Actions repository variables.'
  );
}

/**
 * MSAL configuration.
 *
 * Redirect URI is computed from window.location to support both local dev (http://localhost:5173/)
 * and GitHub Pages production (https://cahp-inc.github.io/cahp-compliance-hub/).
 * Both URIs must be registered on the Azure AD app under Authentication → SPA platform.
 */
export const msalConfig: Configuration = {
  auth: {
    clientId: clientId || '',
    authority: `https://login.microsoftonline.com/${tenantId || 'common'}`,
    redirectUri: window.location.origin + (import.meta.env.BASE_URL || '/'),
    postLogoutRedirectUri: window.location.origin + (import.meta.env.BASE_URL || '/'),
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === LogLevel.Error) console.error('[MSAL]', message);
        else if (import.meta.env.DEV && level === LogLevel.Warning) console.warn('[MSAL]', message);
      },
      logLevel: import.meta.env.DEV ? LogLevel.Warning : LogLevel.Error,
      piiLoggingEnabled: false,
    },
  },
};

/**
 * Scopes requested at sign-in. Includes both User.Read (profile) and SharePoint scopes
 * (admin-consented on the Azure AD app, so they grant silently per-user).
 *
 * Acquiring tokens for these scopes later is handled by `lib/sharepoint/client.ts`
 * via acquireTokenSilent with the same scope set.
 */
export const loginRequest = {
  scopes: ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All'],
};

/**
 * Singleton MSAL instance. Must be initialized before the React tree mounts —
 * see main.tsx for the bootstrap sequence.
 */
export const msalInstance = new PublicClientApplication(msalConfig);
