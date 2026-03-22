import { useEffect, useRef, useState } from 'react';
import { Sheet, SheetContent, SheetHeader, SheetTitle } from '@/components/ui/sheet';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Slider } from '@/components/ui/slider';
import { Switch } from '@/components/ui/switch';
import { Label } from '@/components/ui/label';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { useChartSettingsStore } from '@/store/chartSettingsStore';
import { useSurveyStore } from '@/store/surveyStore';
import { buildPlotlyConfig } from '@/utils/chartBuilder';
import type { ChartType } from '@/types/chartSettings';
import {
  BarChart3, BarChartHorizontal, PieChart, CircleDot,
  Radar, Circle, TrendingDown, Gauge, Box,
  RotateCcw
} from 'lucide-react';

const CHART_TYPE_OPTIONS: { value: ChartType; label: string; icon: React.ElementType }[] = [
  { value: 'bar_vertical', label: 'Barre V.', icon: BarChart3 },
  { value: 'bar_horizontal', label: 'Barre O.', icon: BarChartHorizontal },
  { value: 'bar_3d', label: 'Barre 3D', icon: Box },
  { value: 'pie', label: 'Torta', icon: PieChart },
  { value: 'donut', label: 'Donut', icon: CircleDot },
  { value: 'radar', label: 'Radar', icon: Radar },
  { value: 'bubble', label: 'Bolle', icon: Circle },
  { value: 'waterfall', label: 'Cascata', icon: TrendingDown },
  { value: 'gauge', label: 'Gauge', icon: Gauge },
];

const COLOR_MODES = [
  { value: 'gradient' as const, label: 'Gradiente' },
  { value: 'traffic_light' as const, label: 'Semaforo' },
  { value: 'solid' as const, label: 'Tinta unita' },
  { value: 'brand' as const, label: 'Brand' },
];

const isBarType = (t: ChartType) => t === 'bar_vertical' || t === 'bar_horizontal' || t === 'bar_3d';

interface Props {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function ChartSettingsPanel({ open, onOpenChange }: Props) {
  const { settings, updateSettings, setQuestionOverride, clearQuestionOverride, clearAllOverrides, resetToDefaults, getEffectiveSettings } = useChartSettingsStore();
  const { parsedSurvey, selectedQuestionId } = useSurveyStore();
  const previewRef = useRef<HTMLDivElement>(null);

  // Live preview
  useEffect(() => {
    if (!open || !previewRef.current) return;
    const scaleQuestions = parsedSurvey?.questions.filter(q => q.type === 'scale_1_10_na') || [];
    const previewQuestion = scaleQuestions.find(q => q.id === selectedQuestionId) || scaleQuestions[0];
    if (!previewQuestion || !parsedSurvey) return;
    const analytics = parsedSurvey.scaleAnalytics.get(previewQuestion.id);
    if (!analytics) return;

    const effective = getEffectiveSettings(previewQuestion.id);
    const { data, layout } = buildPlotlyConfig(analytics, previewQuestion, effective);

    import('plotly.js-dist-min').then((Plotly) => {
      Plotly.default.newPlot(previewRef.current!, data, { ...layout, height: 200, margin: { t: 40, r: 20, b: 30, l: 40 } } as any, { displayModeBar: false, responsive: true });
    });

    return () => {
      if (previewRef.current) {
        import('plotly.js-dist-min').then((Plotly) => { Plotly.default.purge(previewRef.current!); });
      }
    };
  }, [open, settings, parsedSurvey, selectedQuestionId]);

  const scaleQuestions = parsedSurvey?.questions.filter(q => q.type === 'scale_1_10_na') || [];

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent className="w-[400px] sm:w-[460px] overflow-y-auto">
        <SheetHeader>
          <SheetTitle>Personalizza Grafici</SheetTitle>
        </SheetHeader>

        {/* Live preview */}
        <div className="mt-4 mb-4">
          <Label className="text-xs text-muted-foreground mb-1 block">Anteprima in tempo reale</Label>
          <div ref={previewRef} className="w-full rounded-lg border border-border bg-background" style={{ height: 200 }} />
        </div>

        <Tabs defaultValue="global" className="mt-2">
          <TabsList className="w-full">
            <TabsTrigger value="global" className="flex-1">Globale</TabsTrigger>
            <TabsTrigger value="per-question" className="flex-1">Per domanda</TabsTrigger>
          </TabsList>

          {/* TAB: Global */}
          <TabsContent value="global" className="space-y-5 mt-4">
            {/* Chart type grid */}
            <div>
              <Label className="text-sm font-medium mb-2 block">Tipo grafico</Label>
              <div className="grid grid-cols-5 gap-1.5">
                {CHART_TYPE_OPTIONS.map(({ value, label, icon: Icon }) => (
                  <button
                    key={value}
                    onClick={() => updateSettings({ chartType: value })}
                    className={`flex flex-col items-center gap-1 p-2 rounded-lg border text-xs transition-all ${
                      settings.chartType === value
                        ? 'border-primary bg-primary/10 text-primary font-medium'
                        : 'border-border hover:border-muted-foreground/40 text-muted-foreground'
                    }`}
                  >
                    <Icon className="w-5 h-5" />
                    <span className="leading-tight text-center">{label}</span>
                  </button>
                ))}
              </div>
            </div>

            {/* Bar spacing */}
            {isBarType(settings.chartType) && (
              <div>
                <Label className="text-sm">Distanza tra barre: {settings.barSpacing.toFixed(2)}</Label>
                <Slider
                  value={[settings.barSpacing]}
                  min={0.1} max={0.9} step={0.05}
                  onValueChange={([v]) => updateSettings({ barSpacing: v })}
                  className="mt-2"
                />
              </div>
            )}

            {/* 3D depth */}
            {settings.chartType === 'bar_3d' && (
              <div>
                <Label className="text-sm">Profondità 3D: {settings.depthEffect}</Label>
                <Slider
                  value={[settings.depthEffect]}
                  min={2} max={20} step={1}
                  onValueChange={([v]) => updateSettings({ depthEffect: v })}
                  className="mt-2"
                />
              </div>
            )}

            {/* Color mode */}
            <div>
              <Label className="text-sm font-medium mb-2 block">Modalità colore</Label>
              <div className="grid grid-cols-2 gap-2">
                {COLOR_MODES.map(({ value, label }) => (
                  <button
                    key={value}
                    onClick={() => updateSettings({ colorMode: value })}
                    className={`px-3 py-2 rounded-lg border text-sm transition-all ${
                      settings.colorMode === value
                        ? 'border-primary bg-primary/10 text-primary font-medium'
                        : 'border-border hover:border-muted-foreground/40 text-muted-foreground'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>

            {/* Gradient colors */}
            {settings.colorMode === 'gradient' && (
              <div className="flex gap-4">
                <div className="flex-1">
                  <Label className="text-xs">Colore basso</Label>
                  <input type="color" value={settings.gradientFrom} onChange={(e) => updateSettings({ gradientFrom: e.target.value })} className="w-full h-8 rounded cursor-pointer mt-1" />
                </div>
                <div className="flex-1">
                  <Label className="text-xs">Colore alto</Label>
                  <input type="color" value={settings.gradientTo} onChange={(e) => updateSettings({ gradientTo: e.target.value })} className="w-full h-8 rounded cursor-pointer mt-1" />
                </div>
              </div>
            )}

            {/* Solid color */}
            {settings.colorMode === 'solid' && (
              <div>
                <Label className="text-xs">Colore barre</Label>
                <input type="color" value={settings.solidColor} onChange={(e) => updateSettings({ solidColor: e.target.value })} className="w-full h-8 rounded cursor-pointer mt-1" />
              </div>
            )}

            {/* Toggles */}
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <Label className="text-sm">Mostra etichette valori</Label>
                <Switch checked={settings.showDataLabels} onCheckedChange={(v) => updateSettings({ showDataLabels: v })} />
              </div>
              {settings.showDataLabels && (
                <div className="flex items-center gap-3 ml-4">
                  <Label className="text-xs text-muted-foreground">Posizione:</Label>
                  <Select value={settings.dataLabelPosition} onValueChange={(v: 'inside' | 'outside') => updateSettings({ dataLabelPosition: v })}>
                    <SelectTrigger className="w-28 h-7 text-xs"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="inside">Interno</SelectItem>
                      <SelectItem value="outside">Esterno</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              )}
              <div className="flex items-center justify-between">
                <Label className="text-sm">Mostra linea media</Label>
                <Switch checked={settings.showMean} onCheckedChange={(v) => updateSettings({ showMean: v })} />
              </div>
              <div className="flex items-center justify-between">
                <Label className="text-sm">Griglia di sfondo</Label>
                <Switch checked={settings.showGridLines} onCheckedChange={(v) => updateSettings({ showGridLines: v })} />
              </div>
            </div>

            {/* Chart height */}
            <div>
              <Label className="text-sm">Altezza grafico</Label>
              <Select value={settings.chartHeight} onValueChange={(v: 'auto' | 'compact' | 'tall') => updateSettings({ chartHeight: v })}>
                <SelectTrigger className="mt-1 h-8 text-sm"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="compact">Compatto</SelectItem>
                  <SelectItem value="auto">Auto</SelectItem>
                  <SelectItem value="tall">Alto</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {/* Reset */}
            <Button variant="outline" size="sm" onClick={resetToDefaults} className="w-full gap-2">
              <RotateCcw className="w-3.5 h-3.5" />
              Ripristina impostazioni predefinite
            </Button>
          </TabsContent>

          {/* TAB: Per question */}
          <TabsContent value="per-question" className="space-y-3 mt-4">
            {scaleQuestions.length === 0 && (
              <p className="text-sm text-muted-foreground py-4 text-center">Nessuna domanda scala caricata.</p>
            )}
            {scaleQuestions.map((q) => {
              const hasOverride = !!settings.perQuestionOverrides[q.id];
              const override = settings.perQuestionOverrides[q.id] || {};
              return (
                <div key={q.id} className="rounded-lg border border-border p-3">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-medium truncate flex-1">
                      {q.questionKey || ''} {q.questionText.slice(0, 50)}
                    </span>
                    <Switch checked={hasOverride} onCheckedChange={(v) => {
                      if (v) setQuestionOverride(q.id, { chartType: settings.chartType });
                      else clearQuestionOverride(q.id);
                    }} />
                  </div>
                  {hasOverride && (
                    <div className="mt-3 space-y-2">
                      <div className="grid grid-cols-5 gap-1">
                        {CHART_TYPE_OPTIONS.map(({ value, label, icon: Icon }) => (
                          <button
                            key={value}
                            onClick={() => setQuestionOverride(q.id, { ...override, chartType: value })}
                            className={`flex flex-col items-center gap-0.5 p-1.5 rounded border text-[10px] ${
                              (override.chartType || settings.chartType) === value
                                ? 'border-primary bg-primary/10 text-primary'
                                : 'border-border text-muted-foreground'
                            }`}
                          >
                            <Icon className="w-3.5 h-3.5" />
                            {label}
                          </button>
                        ))}
                      </div>
                      <div className="grid grid-cols-2 gap-1">
                        {COLOR_MODES.map(({ value, label }) => (
                          <button
                            key={value}
                            onClick={() => setQuestionOverride(q.id, { ...override, colorMode: value })}
                            className={`px-2 py-1 rounded border text-[10px] ${
                              (override.colorMode || settings.colorMode) === value
                                ? 'border-primary bg-primary/10 text-primary'
                                : 'border-border text-muted-foreground'
                            }`}
                          >
                            {label}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
            {scaleQuestions.length > 0 && (
              <Button variant="outline" size="sm" onClick={clearAllOverrides} className="w-full gap-2 mt-2">
                <RotateCcw className="w-3.5 h-3.5" />
                Reset tutti
              </Button>
            )}
          </TabsContent>
        </Tabs>
      </SheetContent>
    </Sheet>
  );
}
