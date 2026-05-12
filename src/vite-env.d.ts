/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_BASE_PATH: string;
  // MSAL config (used starting PR-02)
  readonly VITE_AZURE_CLIENT_ID: string;
  readonly VITE_AZURE_TENANT_ID: string;
  // SharePoint site (used starting PR-03)
  readonly VITE_SHAREPOINT_SITE: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
