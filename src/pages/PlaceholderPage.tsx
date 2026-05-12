import { Icon, IconName } from '../components/ui/Icon';

interface PlaceholderPageProps {
  title: string;
  subtitle?: string;
  icon: IconName;
  plannedPR: string;
}

export function PlaceholderPage({ title, subtitle, icon, plannedPR }: PlaceholderPageProps) {
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
        <p className="text-sm text-gray-500 mb-4">
          This module will be built in <span className="font-mono-data font-semibold text-teal-700">{plannedPR}</span>.
        </p>
        <p className="text-xs text-gray-400 max-w-md mx-auto">
          PR-01 establishes the application shell, navigation, design tokens, and deployment pipeline.
          Individual modules wire to real data starting in PR-03.
        </p>
      </div>
    </div>
  );
}
