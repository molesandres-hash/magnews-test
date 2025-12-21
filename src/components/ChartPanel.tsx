import { useMemo, useEffect, useRef } from 'react';
import { useSurveyStore } from '@/store/surveyStore';
import { BarChart3 } from 'lucide-react';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];
const CHART_COLORS = [
  '#3B82F6', '#3B82F6', '#3B82F6', '#3B82F6', '#3B82F6',
  '#60A5FA', '#60A5FA', '#60A5FA', '#93C5FD', '#BFDBFE', '#94A3B8'
];

export function ChartPanel() {
  const { parsedSurvey, selectedQuestionId } = useSurveyStore();
  const chartRef = useRef<HTMLDivElement>(null);

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
    const values = SCALE_ORDER.map(key => analytics.counts[key] || 0);

    // Dynamic import for Plotly
    import('plotly.js-dist-min').then((Plotly) => {
      const data = [{
        x: SCALE_ORDER,
        y: values,
        type: 'bar' as const,
        marker: {
          color: CHART_COLORS,
          line: { color: '#1E40AF', width: 1 },
        },
        text: values.map(v => v.toString()),
        textposition: 'outside' as const,
      }];

      const layout = {
        title: {
          text: `${question.questionKey || ''} - ${question.questionText.slice(0, 60)}${question.questionText.length > 60 ? '...' : ''}`,
          font: { size: 14, color: '#1E293B' },
        },
        annotations: [{
          x: 0.5,
          y: 1.12,
          xref: 'paper' as const,
          yref: 'paper' as const,
          text: `Media: ${analytics.mean.toFixed(2)} | Risposte: ${analytics.validResponses}/${analytics.totalRespondents}`,
          showarrow: false,
          font: { size: 11, color: '#64748B' },
        }],
        xaxis: { title: { text: 'Valutazione' }, tickfont: { size: 11 } },
        yaxis: { title: { text: 'Conteggio' }, tickfont: { size: 11 } },
        margin: { t: 80, r: 30, b: 50, l: 50 },
        paper_bgcolor: '#FFFFFF',
        plot_bgcolor: '#F8FAFC',
        autosize: true,
        height: 380,
      };

      Plotly.default.newPlot(chartRef.current!, data, layout as any, {
        displayModeBar: false,
        responsive: true,
      });
    });

    return () => {
      if (chartRef.current) {
        import('plotly.js-dist-min').then((Plotly) => {
          Plotly.default.purge(chartRef.current!);
        });
      }
    };
  }, [chartData]);

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
        <div ref={chartRef} style={{ width: '100%', height: '380px' }} />
      </div>
    </div>
  );
}
