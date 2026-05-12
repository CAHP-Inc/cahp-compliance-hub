import { Icon, IconName } from '../components/ui/Icon';
import { Link } from 'react-router-dom';

interface PlaceholderPageProps {
  title: string;
  subtitle?: string;
  icon: IconName;
  plannedPR: string;
}

export function PlaceholderPage({ title, subtitle, icon, plannedPR }: PlaceholderPageProps) {
  const isBacklog = plannedPR === 'Backlog';
  const isPhase = plannedPR.startsWith('Phase');

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-teal-700">{title}</h1>
        {subtitle && <p className="text-sm text-gray-500 mt-1">{subtitle}</p>}
      </div>

      <div className="bg-white border border-gray-200 rounded-lg p-12 text-center shadow-card">
        <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-teal-50 text-teal-700 mb-4">
          <Icon name={icon} size={32} />
        </div>
        <h2 className="text-lg font-semibold text-gray-700 mb-2">{title}</h2>
        {isBacklog ? (
          <>
            <p className="text-sm text-gray-500 mb-2">
              Originally scoped for earlier PRs but not yet built.
            </p>
            <p className="text-xs text-gray-400 max-w-md mx-auto">
              See the <Link to="/" className="text-teal-700 hover:text-teal-900 font-medium">Backlog</Link>{' '}
              section on My Day for the full list of outstanding work and current scheduling status.
            </p>
          </>
        ) : isPhase ? (
          <>
            <p className="text-sm text-gray-500 mb-2">
              Scheduled for <span className="font-mono-data font-semibold text-teal-700">{plannedPR}</span>.
            </p>
            <p className="text-xs text-gray-400 max-w-md mx-auto">
              Phase 1 wraps with PR-07. Phase 2 covers Submittals, Documents, and DOR workflow.
              Phase 3 covers Billing, Reports, and remaining modules.
            </p>
          </>
        ) : (
          <>
            <p className="text-sm text-gray-500 mb-2">
              Scheduled for <span className="font-mono-data font-semibold text-teal-700">{plannedPR}</span>.
            </p>
            <p className="text-xs text-gray-400 max-w-md mx-auto">
              See Phase 1 Build Progress on My Day for the current build order.
            </p>
          </>
        )}
      </div>
    </div>
  );
}
