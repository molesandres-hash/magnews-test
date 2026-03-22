export type ChartType =
  | 'bar_vertical'
  | 'bar_horizontal'
  | 'bar_3d'
  | 'pie'
  | 'donut'
  | 'radar'
  | 'bubble'
  | 'waterfall'
  | 'gauge';

export interface ChartSettings {
  chartType: ChartType;

  // Bar settings
  barSpacing: number;
  barBorderRadius: number;
  barBorderWidth: number;

  // Color settings
  colorMode: 'gradient' | 'solid' | 'traffic_light' | 'brand';
  gradientFrom: string;
  gradientTo: string;
  solidColor: string;

  // 3D settings
  depthEffect: number;

  // Labels
  showDataLabels: boolean;
  dataLabelPosition: 'inside' | 'outside' | 'none';
  showMean: boolean;

  // Layout
  showGridLines: boolean;
  showLegend: boolean;
  chartHeight: 'auto' | 'compact' | 'tall';

  // Per-question overrides
  perQuestionOverrides: Record<string, Partial<ChartSettings>>;
}

export const DEFAULT_CHART_SETTINGS: ChartSettings = {
  chartType: 'bar_vertical',
  barSpacing: 0.3,
  barBorderRadius: 4,
  barBorderWidth: 1,
  colorMode: 'gradient',
  gradientFrom: '#EF4444',
  gradientTo: '#22C55E',
  solidColor: '#3B82F6',
  depthEffect: 8,
  showDataLabels: true,
  dataLabelPosition: 'outside',
  showMean: false,
  showGridLines: true,
  showLegend: false,
  chartHeight: 'auto',
  perQuestionOverrides: {},
};
