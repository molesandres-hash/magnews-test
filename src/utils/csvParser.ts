import Papa from 'papaparse';
import type { ParsedSurvey, Respondent, QuestionInfo, ScaleAnalytics, ClosedAnalytics, OpenAnalytics } from '@/types/survey';
import { normalizeHeader, extractKeyAndText } from './headerNormalizer';
import { classifyQuestion, isNAValue, parseNumericValue } from './questionClassifier';

// Metadata columns to exclude from question detection
const METADATA_COLUMNS = new Set([
  'id contact', 'id database', 'id survey', 'id session', 'test session',
  'status', 'ip address', 'start date', 'end date', 'data inizio', 'data fine',
  'indirizzo ip', 'sessione di test', 'stato'
]);

/**
 * Detect if a column is a metadata column
 */
function isMetadataColumn(header: string): boolean {
  const lower = header.toLowerCase().trim();
  if (METADATA_COLUMNS.has(lower)) return true;
  if (lower.startsWith('page ')) return true;
  return false;
}

/**
 * Parse a CSV file and return structured survey data
 */
export async function parseCSVFile(file: File): Promise<ParsedSurvey> {
  return new Promise((resolve, reject) => {
    const warnings: string[] = [];

    Papa.parse(file, {
      complete: (result) => {
        try {
          const survey = processParsedData(result.data as string[][], file.name, warnings);
          resolve(survey);
        } catch (error) {
          reject(error);
        }
      },
      error: (error) => {
        reject(new Error(`CSV parsing error: ${error.message}`));
      },
      skipEmptyLines: true,
      encoding: 'UTF-8',
    });
  });
}

/**
 * Process parsed CSV data into structured survey format
 */
function processParsedData(
  rawData: string[][],
  fileName: string,
  warnings: string[]
): ParsedSurvey {
  if (rawData.length < 2) {
    throw new Error('Il file CSV deve contenere almeno una riga di intestazione e una riga di dati.');
  }

  const headers = rawData[0];
  const dataRows = rawData.slice(1);

  // Find column indices
  const columnIndices = findColumnIndices(headers);

  // Process respondents
  const { respondents, completedCount, excludedCount, testSessionCount } = processRespondents(
    dataRows,
    headers,
    columnIndices,
    warnings
  );

  if (completedCount === 0) {
    throw new Error('Nessuna risposta completata trovata nel file. Verificare la colonna "Status".');
  }

  // Identify and process questions
  const questions = identifyQuestions(headers, respondents, warnings);

  // Compute analytics
  const scaleAnalytics = computeScaleAnalytics(questions, respondents);
  const closedAnalytics = computeClosedAnalytics(questions, respondents);
  const openAnalytics = computeOpenAnalytics(questions, respondents);

  return {
    rawData,
    headers,
    respondents,
    questions,
    scaleAnalytics,
    closedAnalytics,
    openAnalytics,
    metadata: {
      totalRows: dataRows.length,
      completedCount,
      excludedCount,
      testSessionCount,
      warnings,
      parsedAt: new Date(),
      fileName,
    },
  };
}

/**
 * Find indices of important columns
 */
function findColumnIndices(headers: string[]): Record<string, number> {
  const indices: Record<string, number> = {};
  
  headers.forEach((h, i) => {
    const lower = h.toLowerCase().trim();
    if (lower === 'status' || lower === 'stato') indices.status = i;
    if (lower === 'end date' || lower === 'data fine') indices.endDate = i;
    if (lower === 'start date' || lower === 'data inizio') indices.startDate = i;
    if (lower === 'cognome') indices.cognome = i;
    if (lower === 'nome') indices.nome = i;
    if (lower === 'id contact') indices.idContact = i;
    if (lower === 'id session') indices.idSession = i;
    if (lower === 'test session' || lower === 'sessione di test') indices.testSession = i;
  });

  return indices;
}

/**
 * Process respondents and filter by completion status
 */
function processRespondents(
  dataRows: string[][],
  headers: string[],
  columnIndices: Record<string, number>,
  warnings: string[]
): {
  respondents: Respondent[];
  completedCount: number;
  excludedCount: number;
  testSessionCount: number;
} {
  const nameTracker = new Map<string, number>();
  const respondents: Respondent[] = [];
  let completedCount = 0;
  let excludedCount = 0;
  let testSessionCount = 0;

  for (const row of dataRows) {
    // Build original data map
    const originalData: Record<string, string> = {};
    headers.forEach((h, i) => {
      originalData[h] = row[i] || '';
    });

    // Check test session
    const isTestSession = columnIndices.testSession !== undefined && 
      row[columnIndices.testSession]?.toLowerCase() === 'yes';
    
    if (isTestSession) {
      testSessionCount++;
      continue; // Skip test sessions
    }

    // Determine completion status
    let status = '';
    let isCompleted = false;

    if (columnIndices.status !== undefined) {
      status = row[columnIndices.status] || '';
      isCompleted = status.toLowerCase() === 'completed' || status.toLowerCase() === 'completato';
    } else if (columnIndices.endDate !== undefined) {
      isCompleted = !!row[columnIndices.endDate]?.trim();
      if (!warnings.includes('No "Status" column found')) {
        warnings.push('Colonna "Status" non trovata, usando "End date" per determinare le risposte complete.');
      }
    } else if (columnIndices.startDate !== undefined) {
      isCompleted = !!row[columnIndices.startDate]?.trim();
      if (!warnings.includes('No status or end date')) {
        warnings.push('Colonne "Status" e "End date" non trovate, usando "Start date" per determinare le risposte.');
      }
    } else {
      isCompleted = true;
      if (!warnings.includes('No status columns')) {
        warnings.push('Nessuna colonna di stato trovata, includendo tutte le righe.');
      }
    }

    if (!isCompleted) {
      excludedCount++;
      continue;
    }

    completedCount++;

    // Generate display name
    let baseName = '';
    if (columnIndices.cognome !== undefined && columnIndices.nome !== undefined) {
      const cognome = row[columnIndices.cognome]?.trim() || '';
      const nome = row[columnIndices.nome]?.trim() || '';
      baseName = `${cognome} ${nome}`.trim();
    }
    if (!baseName && columnIndices.idContact !== undefined) {
      baseName = row[columnIndices.idContact]?.trim() || '';
    }
    if (!baseName && columnIndices.idSession !== undefined) {
      baseName = row[columnIndices.idSession]?.trim() || '';
    }
    if (!baseName) {
      baseName = `Rispondente ${completedCount}`;
    }

    // Ensure uniqueness
    const count = nameTracker.get(baseName) || 0;
    nameTracker.set(baseName, count + 1);
    const displayName = count > 0 ? `${baseName} (${count + 1})` : baseName;

    respondents.push({
      id: `resp_${respondents.length}`,
      displayName,
      originalData,
      status,
      isCompleted,
      isTestSession,
    });
  }

  return { respondents, completedCount, excludedCount, testSessionCount };
}

/**
 * Identify question columns and classify them
 */
function identifyQuestions(
  headers: string[],
  respondents: Respondent[],
  warnings: string[]
): QuestionInfo[] {
  const questions: QuestionInfo[] = [];
  const processedKeys = new Set<string>();

  // Group headers by their cleaned base (without values/labels suffix)
  const headerGroups = new Map<string, { valuesIdx?: number; labelsIdx?: number; rawHeader: string }>();

  headers.forEach((header, idx) => {
    if (isMetadataColumn(header)) return;

    const normalized = normalizeHeader(header);
    const baseKey = normalized.cleanedHeader;

    if (!headerGroups.has(baseKey)) {
      headerGroups.set(baseKey, { rawHeader: header });
    }

    const group = headerGroups.get(baseKey)!;
    if (normalized.valueSource === 'values') {
      group.valuesIdx = idx;
    } else if (normalized.valueSource === 'labels') {
      group.labelsIdx = idx;
    } else {
      // If no suffix, treat as values column
      if (group.valuesIdx === undefined) {
        group.valuesIdx = idx;
      }
    }
  });

  // Process each unique question
  headerGroups.forEach((group, baseKey) => {
    // Skip if already processed or if it's just a labels column without values
    if (processedKeys.has(baseKey)) return;
    
    const columnIdx = group.valuesIdx ?? group.labelsIdx;
    if (columnIdx === undefined) return;

    processedKeys.add(baseKey);

    const normalized = normalizeHeader(headers[columnIdx]);
    const { questionKey, questionText, blockId, subId } = extractKeyAndText(normalized.cleanedHeader);

    // Get values for classification
    const values = respondents.map(r => r.originalData[headers[columnIdx]] || '');
    const valueSource = group.valuesIdx !== undefined ? 'values' : 'labels';
    const type = classifyQuestion(values, valueSource as 'values' | 'labels' | 'unknown');

    if (type === 'unknown') {
      warnings.push(`Tipo di domanda non determinato per: "${normalized.cleanedHeader.slice(0, 50)}..."`);
    }

    questions.push({
      id: `q_${questions.length}`,
      rawHeader: headers[columnIdx],
      cleanedHeader: normalized.cleanedHeader,
      questionKey,
      questionText,
      blockId,
      subId,
      type,
      valueSource: valueSource as 'values' | 'labels' | 'unknown',
      valuesColumnIndex: columnIdx,
      labelsColumnIndex: group.labelsIdx,
    });
  });

  // Sort questions by block and sub ID
  questions.sort((a, b) => {
    // Null blocks go last
    if (a.blockId === null && b.blockId !== null) return 1;
    if (a.blockId !== null && b.blockId === null) return -1;
    if (a.blockId === null && b.blockId === null) return 0;
    
    if (a.blockId! !== b.blockId!) {
      return a.blockId! - b.blockId!;
    }
    return a.subId - b.subId;
  });

  return questions;
}

/**
 * Compute analytics for scale questions
 */
function computeScaleAnalytics(
  questions: QuestionInfo[],
  respondents: Respondent[]
): Map<string, ScaleAnalytics> {
  const analytics = new Map<string, ScaleAnalytics>();

  questions
    .filter(q => q.type === 'scale_1_10_na')
    .forEach(question => {
      const counts: Record<string, number> = {
        '10': 0, '9': 0, '8': 0, '7': 0, '6': 0,
        '5': 0, '4': 0, '3': 0, '2': 0, '1': 0, 'N/A': 0
      };
      
      const respondentValues: Record<string, number | null> = {};
      let sum = 0;
      let validCount = 0;

      respondents.forEach(respondent => {
        const rawValue = respondent.originalData[question.rawHeader];
        const numValue = parseNumericValue(rawValue || '');

        if (numValue !== null && numValue >= 1 && numValue <= 10) {
          const key = Math.round(numValue).toString();
          counts[key] = (counts[key] || 0) + 1;
          respondentValues[respondent.id] = numValue;
          sum += numValue;
          validCount++;
        } else {
          counts['N/A']++;
          respondentValues[respondent.id] = null;
        }
      });

      const mean = validCount > 0 ? Math.round((sum / validCount) * 100) / 100 : 0;

      analytics.set(question.id, {
        questionId: question.id,
        mean,
        counts,
        totalRespondents: respondents.length,
        validResponses: validCount,
        respondentValues,
      });
    });

  return analytics;
}

/**
 * Compute analytics for closed questions
 */
function computeClosedAnalytics(
  questions: QuestionInfo[],
  respondents: Respondent[]
): Map<string, ClosedAnalytics> {
  const analytics = new Map<string, ClosedAnalytics>();

  questions
    .filter(q => q.type === 'closed_single' || q.type === 'closed_binary' || q.type === 'closed_multi')
    .forEach(question => {
      const optionCounts = new Map<string, number>();
      let totalResponses = 0;

      respondents.forEach(respondent => {
        const rawValue = respondent.originalData[question.rawHeader];
        if (isNAValue(rawValue)) return;

        totalResponses++;

        if (question.type === 'closed_multi') {
          // Split multi-select values
          const options = rawValue.split(/[;|]/).map(o => o.trim()).filter(Boolean);
          options.forEach(opt => {
            optionCounts.set(opt, (optionCounts.get(opt) || 0) + 1);
          });
        } else {
          optionCounts.set(rawValue.trim(), (optionCounts.get(rawValue.trim()) || 0) + 1);
        }
      });

      const options = Array.from(optionCounts.entries())
        .map(([option, count]) => ({
          option,
          count,
          percent: totalResponses > 0 ? Math.round((count / totalResponses) * 1000) / 10 : 0,
        }))
        .sort((a, b) => b.count - a.count);

      analytics.set(question.id, {
        questionId: question.id,
        options,
        totalRespondents: respondents.length,
      });
    });

  return analytics;
}

/**
 * Compute analytics for open text questions
 */
function computeOpenAnalytics(
  questions: QuestionInfo[],
  respondents: Respondent[]
): Map<string, OpenAnalytics> {
  const analytics = new Map<string, OpenAnalytics>();

  questions
    .filter(q => q.type === 'open_text')
    .forEach(question => {
      const responses: Array<{ respondentId: string; respondentName: string; answer: string }> = [];
      let filledCount = 0;
      let emptyCount = 0;

      respondents.forEach(respondent => {
        const rawValue = respondent.originalData[question.rawHeader];
        
        if (isNAValue(rawValue)) {
          emptyCount++;
        } else {
          filledCount++;
          responses.push({
            respondentId: respondent.id,
            respondentName: respondent.displayName,
            answer: rawValue.trim(),
          });
        }
      });

      analytics.set(question.id, {
        questionId: question.id,
        responses,
        filledCount,
        emptyCount,
      });
    });

  return analytics;
}
