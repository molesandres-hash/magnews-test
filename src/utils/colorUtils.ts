import type { ChartSettings } from '@/types/chartSettings';
import { useTemplateStore } from '@/store/templateStore';

const SCALE = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result
    ? { r: parseInt(result[1], 16), g: parseInt(result[2], 16), b: parseInt(result[3], 16) }
    : { r: 59, g: 130, b: 246 };
}

export function hexToRgba(hex: string, alpha: number): string {
  const { r, g, b } = hexToRgb(hex);
  return `rgba(${r},${g},${b},${alpha})`;
}

export function interpolateColor(from: string, to: string, t: number): string {
  const f = hexToRgb(from);
  const toRgb = hexToRgb(to);
  const r = Math.round(f.r + (toRgb.r - f.r) * t);
  const g = Math.round(f.g + (toRgb.g - f.g) * t);
  const b = Math.round(f.b + (toRgb.b - f.b) * t);
  return `rgb(${r},${g},${b})`;
}

export function darkenHex(hex: string, amount: number): string {
  const { r, g, b } = hexToRgb(hex);
  return `rgb(${Math.max(0, r - amount)},${Math.max(0, g - amount)},${Math.max(0, b - amount)})`;
}

export function lightenHex(hex: string, amount: number): string {
  const { r, g, b } = hexToRgb(hex);
  return `rgb(${Math.min(255, r + amount)},${Math.min(255, g + amount)},${Math.min(255, b + amount)})`;
}

export function getColorArray(settings: ChartSettings): string[] {
  if (settings.colorMode === 'traffic_light') {
    return SCALE.map((s) => {
      const n = parseInt(s);
      if (isNaN(n)) return '#94A3B8';
      if (n >= 8) return '#22C55E';
      if (n >= 5) return '#F59E0B';
      return '#EF4444';
    });
  }

  if (settings.colorMode === 'solid') {
    return SCALE.map((s) => (s === 'N/A' ? '#94A3B8' : settings.solidColor));
  }

  if (settings.colorMode === 'brand') {
    const template = useTemplateStore.getState().getActiveTemplate();
    if (template) {
      return SCALE.map((s, i) => {
        if (s === 'N/A') return '#94A3B8';
        return interpolateColor(template.primaryColor, template.secondaryColor, i / 9);
      });
    }
  }

  // gradient (default)
  return SCALE.map((s, i) => {
    if (s === 'N/A') return '#94A3B8';
    return interpolateColor(settings.gradientFrom, settings.gradientTo, 1 - i / 9);
  });
}
