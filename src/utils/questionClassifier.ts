import type { QuestionType } from '@/types/survey';

// Tokens that represent N/A or missing values
const NA_TOKENS = new Set([
  '', 'n/a', 'na', 'n.a.', 'n.a', '-', '--', 'null', 'none', 
  'nessuna risposta', 'no response', 'non applicabile', 'not applicable'
]);

/**
 * Normalize a value to check if it's an N/A token
 */
export function isNAValue(value: string | null | undefined): boolean {
  if (value === null || value === undefined) return true;
  const normalized = value.toString().toLowerCase().trim();
  return NA_TOKENS.has(normalized);
}

/**
 * Try to parse a value as a number
 */
export function parseNumericValue(value: string): number | null {
  if (isNAValue(value)) return null;
  
  let cleaned = value.trim();
  
  // Strip "v" or "V" prefix commonly used in Magnews scale values (e.g., "v10" -> "10")
  if (/^[vV]\d/.test(cleaned)) {
    cleaned = cleaned.substring(1);
  }
  
  cleaned = cleaned.replace(',', '.');
  const num = parseFloat(cleaned);
  
  if (isNaN(num)) return null;
  return num;
}

/**
 * Classify a question based on its values
 */
export function classifyQuestion(
  values: string[],
  valueSource: 'values' | 'labels' | 'unknown'
): QuestionType {
  // Filter out empty values
  const nonEmptyValues = values.filter(v => !isNAValue(v));
  
  if (nonEmptyValues.length === 0) {
    return 'unknown';
  }

  // Get unique values
  const uniqueValues = [...new Set(nonEmptyValues.map(v => v.trim().toLowerCase()))];

  // Check if it's a scale 1-10 question
  if (valueSource === 'values' || valueSource === 'unknown') {
    const numericValues = nonEmptyValues
      .map(parseNumericValue)
      .filter((n): n is number => n !== null);
    
    if (numericValues.length > 0) {
      const uniqueNumeric = [...new Set(numericValues)];
      const allInScale = uniqueNumeric.every(n => n >= 1 && n <= 10 && Number.isInteger(n));
      
      if (allInScale && uniqueNumeric.length > 1) {
        return 'scale_1_10_na';
      }
    }
  }

  // Check for yes/no binary questions
  const yesNoTokens = new Set(['sì', 'si', 'yes', 'no', 'vero', 'falso', 'true', 'false', '1', '0']);
  if (uniqueValues.length <= 2 && uniqueValues.every(v => yesNoTokens.has(v))) {
    return 'closed_binary';
  }

  // Check for open text (high uniqueness or long average length)
  const avgLength = nonEmptyValues.reduce((sum, v) => sum + v.length, 0) / nonEmptyValues.length;
  const uniquenessRatio = uniqueValues.length / nonEmptyValues.length;

  if (avgLength > 50 || (uniquenessRatio > 0.5 && uniqueValues.length > 5)) {
    return 'open_text';
  }

  // Check for multi-select (contains separators)
  const hasMultiSeparators = nonEmptyValues.some(v => 
    v.includes(';') || v.includes('|') || (v.includes(',') && v.split(',').length > 2)
  );

  if (hasMultiSeparators) {
    return 'closed_multi';
  }

  // Default to closed single if small unique set
  if (uniqueValues.length <= 15) {
    return 'closed_single';
  }

  return 'unknown';
}

/**
 * Get human-readable label for question type
 */
export function getQuestionTypeLabel(type: QuestionType): string {
  switch (type) {
    case 'scale_1_10_na':
      return 'Scala 1-10';
    case 'open_text':
      return 'Risposta Aperta';
    case 'closed_single':
      return 'Scelta Singola';
    case 'closed_binary':
      return 'Sì/No';
    case 'closed_multi':
      return 'Scelta Multipla';
    case 'unknown':
      return 'Tipo Sconosciuto';
  }
}

/**
 * Get color class for question type badge
 */
export function getQuestionTypeColor(type: QuestionType): string {
  switch (type) {
    case 'scale_1_10_na':
      return 'bg-chart-1/10 text-chart-1 border-chart-1/20';
    case 'open_text':
      return 'bg-chart-3/10 text-chart-3 border-chart-3/20';
    case 'closed_single':
      return 'bg-chart-2/10 text-chart-2 border-chart-2/20';
    case 'closed_binary':
      return 'bg-chart-4/10 text-chart-4 border-chart-4/20';
    case 'closed_multi':
      return 'bg-chart-6/10 text-chart-6 border-chart-6/20';
    case 'unknown':
      return 'bg-muted text-muted-foreground border-muted';
  }
}
