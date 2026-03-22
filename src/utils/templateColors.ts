import type { CompanyTemplate } from '@/types/companyTemplate';

/**
 * Interpolate between two hex colors
 */
function interpolateColor(color1: string, color2: string, factor: number): string {
  const r1 = parseInt(color1.slice(1, 3), 16);
  const g1 = parseInt(color1.slice(3, 5), 16);
  const b1 = parseInt(color1.slice(5, 7), 16);
  const r2 = parseInt(color2.slice(1, 3), 16);
  const g2 = parseInt(color2.slice(3, 5), 16);
  const b2 = parseInt(color2.slice(5, 7), 16);
  const r = Math.round(r1 + (r2 - r1) * factor);
  const g = Math.round(g1 + (g2 - g1) * factor);
  const b = Math.round(b1 + (b2 - b1) * factor);
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

/**
 * Generate 11 gradient colors for scale values 10→1 + N/A from template colors
 */
export function getTemplateChartColors(template: CompanyTemplate | null): string[] {
  if (!template) {
    return [
      '#22C55E', '#22C55E', '#22C55E', '#84CC16',
      '#FDE047', '#F59E0B', '#F97316', '#EF4444',
      '#DC2626', '#B91C1C', '#94A3B8'
    ];
  }
  const colors: string[] = [];
  for (let i = 0; i < 10; i++) {
    colors.push(interpolateColor(template.primaryColor, template.secondaryColor, i / 9));
  }
  colors.push('#94A3B8'); // N/A always gray
  return colors;
}

/**
 * Get a light version of secondaryColor for plot background
 */
export function getTemplatePlotBg(template: CompanyTemplate | null): string {
  if (!template) return '#F8FAFC';
  return interpolateColor(template.secondaryColor, '#FFFFFF', 0.8);
}

/**
 * Get bar color based on mean value, with template override
 */
export function getMeanBarColor(mean: number, template: CompanyTemplate | null): string {
  if (template) return template.primaryColor;
  if (mean >= 8) return '#22C55E';
  if (mean >= 6) return '#3B82F6';
  if (mean >= 4) return '#F59E0B';
  return '#EF4444';
}

/**
 * Hex to ARGB (for ExcelJS) - strips # and prepends FF
 */
export function hexToArgb(hex: string): string {
  return 'FF' + hex.replace('#', '').toUpperCase();
}
