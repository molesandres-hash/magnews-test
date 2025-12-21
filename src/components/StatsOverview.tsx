import { useMemo } from 'react';
import { Users, FileQuestion, BarChart3, AlertTriangle, CheckCircle2 } from 'lucide-react';
import { useSurveyStore } from '@/store/surveyStore';
import { getSurveyStats, getUniqueBlocks } from '@/utils/analytics';
import { cn } from '@/lib/utils';

interface StatCardProps {
  label: string;
  value: string | number;
  icon: React.ReactNode;
  variant?: 'default' | 'success' | 'warning';
  subtitle?: string;
}

function StatCard({ label, value, icon, variant = 'default', subtitle }: StatCardProps) {
  return (
    <div className="stat-card animate-scale-in">
      <div className="flex items-start justify-between">
        <div>
          <p className="text-sm font-medium text-muted-foreground mb-1">{label}</p>
          <p className={cn(
            'text-3xl font-bold',
            variant === 'success' && 'text-success',
            variant === 'warning' && 'text-warning',
            variant === 'default' && 'text-foreground'
          )}>
            {value}
          </p>
          {subtitle && (
            <p className="text-xs text-muted-foreground mt-1">{subtitle}</p>
          )}
        </div>
        <div className={cn(
          'w-12 h-12 rounded-xl flex items-center justify-center',
          variant === 'success' && 'bg-success/10 text-success',
          variant === 'warning' && 'bg-warning/10 text-warning',
          variant === 'default' && 'bg-primary/10 text-primary'
        )}>
          {icon}
        </div>
      </div>
    </div>
  );
}

export function StatsOverview() {
  const { parsedSurvey } = useSurveyStore();

  const stats = useMemo(() => {
    if (!parsedSurvey) return null;
    return getSurveyStats(parsedSurvey);
  }, [parsedSurvey]);

  if (!parsedSurvey || !stats) return null;

  const { metadata } = parsedSurvey;

  return (
    <div className="grid grid-cols-2 lg:grid-cols-4 xl:grid-cols-5 gap-4 animate-fade-in">
      <StatCard
        label="Risposte Complete"
        value={metadata.completedCount}
        icon={<Users className="w-6 h-6" />}
        variant="success"
        subtitle={`${metadata.excludedCount} escluse`}
      />
      
      <StatCard
        label="Domande Totali"
        value={stats.totalQuestions}
        icon={<FileQuestion className="w-6 h-6" />}
      />
      
      <StatCard
        label="Domande Scala"
        value={stats.scaleQuestions}
        icon={<BarChart3 className="w-6 h-6" />}
        subtitle={`${stats.openQuestions} aperte, ${stats.closedQuestions} chiuse`}
      />
      
      <StatCard
        label="Blocchi"
        value={stats.blocks}
        icon={<CheckCircle2 className="w-6 h-6" />}
      />
      
      <StatCard
        label="Avvisi"
        value={metadata.warnings.length}
        icon={<AlertTriangle className="w-6 h-6" />}
        variant={metadata.warnings.length > 0 ? 'warning' : 'default'}
      />
    </div>
  );
}
