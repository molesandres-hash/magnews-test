import { useMemo } from 'react';
import Plot from 'react-plotly.js';
import { useSurveyStore } from '@/store/surveyStore';
import { getDistributionChartData } from '@/utils/chartExporter';
import { BarChart3 } from 'lucide-react';

export function ChartPanel() {
  const { parsedSurvey, selectedQuestionId } = useSurveyStore();

  const chartData = useMemo(() => {
    if (!parsedSurvey || !selectedQuestionId) return null;

    const question = parsedSurvey.questions.find(q => q.id === selectedQuestionId);
    if (!question || question.type !== 'scale_1_10_na') return null;

    const analytics = parsedSurvey.scaleAnalytics.get(question.id);
    if (!analytics) return null;

    return getDistributionChartData(analytics, question);
  }, [parsedSurvey, selectedQuestionId]);

  if (!parsedSurvey) return null;

  if (!chartData) {
    return (
      <div className="glass-card rounded-xl p-8 h-[500px] flex flex-col items-center justify-center text-center animate-fade-in">
        <div className="w-16 h-16 rounded-2xl bg-muted flex items-center justify-center mb-4">
          <BarChart3 className="w-8 h-8 text-muted-foreground" />
        </div>
        <h3 className="text-lg font-semibold text-foreground mb-2">
          Seleziona una domanda scala
        </h3>
        <p className="text-sm text-muted-foreground max-w-sm">
          Clicca su una domanda di tipo "Scala 1-10" nella tabella per visualizzare il grafico della distribuzione.
        </p>
      </div>
    );
  }

  return (
    <div className="glass-card rounded-xl overflow-hidden animate-scale-in">
      <div className="p-4 border-b border-border bg-muted/30">
        <h3 className="font-semibold text-lg">Distribuzione Risposte</h3>
      </div>
      <div className="p-4">
        <Plot
          data={chartData.data as any}
          layout={{
            ...chartData.layout,
            autosize: true,
            height: 400,
          } as any}
          config={{
            displayModeBar: false,
            responsive: true,
          }}
          style={{ width: '100%' }}
        />
      </div>
    </div>
  );
}
