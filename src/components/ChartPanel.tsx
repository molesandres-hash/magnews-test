import { useMemo, useEffect, useRef } from 'react';
import { useSurveyStore } from '@/store/surveyStore';
import { useChartSettingsStore } from '@/store/chartSettingsStore';
import { buildPlotlyConfig } from '@/utils/chartBuilder';
import { BarChart3, Palette } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { ChartSettingsPanel } from '@/components/ChartSettingsPanel';
import { useState } from 'react';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

export function ChartPanel() {
  const { parsedSurvey, selectedQuestionId } = useSurveyStore();
  const { settings, getEffectiveSettings } = useChartSettingsStore();
  const chartRef = useRef<HTMLDivElement>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const chartData = useMemo(() => {
    if (!parsedSurvey || !selectedQuestionId) return null;
    const question = parsedSurvey.questions.find(q => q.id === selectedQuestionId);
    if (!question || question.type !== 'scale_1_10_na') return null;
    const analytics = parsedSurvey.scaleAnalytics.get(question.id);
    if (!analytics) return null;
    return { analytics, question };
  }, [parsedSurvey, selectedQuestionId]);

  useEffect(() => {
    if (!chartData || !chartRef.current) return;

    const { analytics, question } = chartData;
    const effective = getEffectiveSettings(question.id);
    const { data, layout } = buildPlotlyConfig(analytics, question, effective);

    import('plotly.js-dist-min').then((Plotly) => {
      Plotly.default.newPlot(chartRef.current!, data, layout as any, { displayModeBar: false, responsive: true });
    });

    return () => {
      if (chartRef.current) {
        import('plotly.js-dist-min').then((Plotly) => {
          Plotly.default.purge(chartRef.current!);
        });
      }
    };
  }, [chartData, settings]);

  if (!parsedSurvey) return null;

  if (!chartData) {
    return (
      <div className="glass-card rounded-xl p-8 h-[500px] flex flex-col items-center justify-center text-center animate-fade-in">
        <div className="w-16 h-16 rounded-2xl bg-muted flex items-center justify-center mb-4">
          <BarChart3 className="w-8 h-8 text-muted-foreground" />
        </div>
        <h3 className="text-lg font-semibold text-foreground mb-2">Seleziona una domanda scala</h3>
        <p className="text-sm text-muted-foreground max-w-sm">
          Clicca su una domanda di tipo "Scala 1-10" nella tabella per visualizzare il grafico della distribuzione.
        </p>
      </div>
    );
  }

  return (
    <div className="glass-card rounded-xl overflow-hidden animate-scale-in relative">
      <div className="p-4 border-b border-border bg-muted/30 flex items-center justify-between">
        <h3 className="font-semibold text-lg">Distribuzione Risposte</h3>
        <Button variant="ghost" size="sm" onClick={() => setSettingsOpen(true)} className="gap-1.5">
          <Palette className="w-4 h-4" />
          <span className="hidden sm:inline text-xs">Personalizza</span>
        </Button>
      </div>
      <div className="p-4">
        <div ref={chartRef} style={{ width: '100%', height: settings.chartHeight === 'compact' ? '280px' : settings.chartHeight === 'tall' ? '520px' : '380px' }} />
      </div>

      <ChartSettingsPanel open={settingsOpen} onOpenChange={setSettingsOpen} />
    </div>
  );
}
