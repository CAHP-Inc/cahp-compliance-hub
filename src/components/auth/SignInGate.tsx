import { ReactNode } from 'react';
import { useMsal } from '@azure/msal-react';
import { useSession } from '../../lib/session';
import { loginRequest } from '../../lib/auth/msalConfig';
import { Icon } from '../ui/Icon';

interface SignInGateProps {
  children: ReactNode;
}

/**
 * Gates the application behind M365 authentication and role-based authorization.
 *
 * Three states:
 *   1. Not authenticated → branded sign-in screen with "Sign in with Microsoft" button
 *   2. Authenticated but not on role map → "Access denied" screen with sign-out option
 *   3. Authenticated and authorized → renders children (the app)
 */
export function SignInGate({ children }: SignInGateProps) {
  const { instance } = useMsal();
  const session = useSession();

  const handleSignIn = () => {
    instance.loginRedirect(loginRequest).catch((err) => {
      console.error('Sign-in failed:', err);
    });
  };

  const handleSignOut = () => {
    instance.logoutRedirect().catch((err) => {
      console.error('Sign-out failed:', err);
    });
  };

  // State 1: not authenticated
  if (!session.isAuthenticated) {
    return (
      <div className="min-h-screen bg-teal-900 flex items-center justify-center p-6">
        <div className="bg-white rounded-xl shadow-2xl p-10 max-w-md w-full text-center">
          <div className="w-16 h-16 rounded-xl bg-teal-900 text-gold-500 font-bold text-3xl flex items-center justify-center mx-auto mb-6">
            C
          </div>
          <h1 className="text-2xl font-bold text-teal-900 mb-2">CAHP Compliance Hub</h1>
          <p className="text-sm text-gray-500 mb-1">Carolina Affordable Housing Project</p>
          <p className="text-sm text-gray-600 mt-8 mb-6">
            Sign in with your Microsoft 365 account to access the application.
          </p>
          <button
            onClick={handleSignIn}
            className="w-full bg-teal-900 hover:bg-teal-700 text-white font-semibold py-3 px-6 rounded-lg transition-colors flex items-center justify-center gap-2"
          >
            <Icon name="external" size={18} />
            Sign in with Microsoft
          </button>
          <p className="text-xs text-gray-400 mt-8">
            Authored by Brandy Turner · NewShire Property Management
          </p>
        </div>
      </div>
    );
  }

  // Loading state: signed in, waiting on the Access List fetch to resolve
  // before deciding access. Without this we'd briefly flash the "Access denied"
  // screen against the fallback list before the SharePoint copy lands.
  if (session.isAuthenticated && session.accessLoading && !session.isAuthorized) {
    return (
      <div className="min-h-screen bg-teal-900 flex items-center justify-center p-6">
        <div className="bg-white rounded-xl shadow-2xl p-10 max-w-md w-full text-center">
          <div className="inline-flex items-center gap-3 text-gray-600">
            <div className="w-4 h-4 rounded-full border-2 border-teal-700 border-r-transparent animate-spin" />
            <span className="text-sm">Checking access…</span>
          </div>
        </div>
      </div>
    );
  }

  // State 2: authenticated but not authorized
  if (!session.isAuthorized) {
    return (
      <div className="min-h-screen bg-teal-900 flex items-center justify-center p-6">
        <div className="bg-white rounded-xl shadow-2xl p-10 max-w-md w-full text-center">
          <div className="w-16 h-16 rounded-full bg-red-50 text-error flex items-center justify-center mx-auto mb-6">
            <Icon name="alert" size={32} />
          </div>
          <h1 className="text-xl font-bold text-teal-900 mb-3">Access denied</h1>
          <p className="text-sm text-gray-600 mb-2">
            Your account is not on the access list for CAHP Compliance Hub.
          </p>
          <p className="text-sm text-gray-600 mb-1">Signed in as:</p>
          <p className="text-sm font-semibold text-gray-800 mb-6">{session.user?.email}</p>
          <p className="text-xs text-gray-500 mb-6">
            Contact Brandy Turner to request access. If you're an admin, open
            <span className="font-mono-data mx-1 px-1.5 py-0.5 bg-gray-100 rounded text-[11px]">
              Settings → Access List
            </span>
            and add this email.
          </p>
          <button
            onClick={handleSignOut}
            className="text-sm text-teal-700 hover:text-teal-900 font-medium underline"
          >
            Sign out and try a different account
          </button>
        </div>
      </div>
    );
  }

  // State 3: authenticated and authorized
  return <>{children}</>;
}
