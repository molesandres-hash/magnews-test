import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

/**
 * Generate file_per_questionari_GENERATO.xlsx
 * Contains 5 sheets: Export, Foglio2, Persone, estrazione per grafici , per pdf 
 * 
 * CRITICAL: Export/Foglio2 contain VALUES
 * estrazione per grafici and per pdf contain FORMULAS referencing Export
 */
export async function generateFilePerQuestionari(survey: ParsedSurvey): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  // Build question row mapping for formula references
  const questionRowMap = new Map<string, number>();

  // Sheet 1: Export - Full transposed matrix (VALUES)
  createExportSheet(workbook, survey, questionRowMap);

  // Sheet 2: Foglio2 - Values-only copy
  createFoglio2Sheet(workbook, survey);

  // Sheet 3: Persone - Respondent metadata
  createPersoneSheet(workbook, survey);

  // Sheet 4: estrazione per grafici  (FORMULAS referencing Export)
  createEstrazioneGraficiSheet(workbook, survey, questionRowMap);

  // Sheet 5: per pdf  (FORMULAS referencing estrazione per grafici)
  createPerPdfSheet(workbook, survey, questionRowMap);

  // Generate and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  saveAs(blob, 'file_per_questionari_GENERATO.xlsx');
}

/**
 * Get respondent's answer for a question (raw value)
 */
function getRespondentAnswer(
  survey: ParsedSurvey,
  question: QuestionInfo,
  respondentId: string
): string | number {
  switch (question.type) {
    case 'scale_1_10_na': {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (analytics) {
        const value = analytics.respondentValues[respondentId];
        return value !== null ? value : 'N/A';
      }
      return '';
    }
    case 'open_text': {
      const analytics = survey.openAnalytics.get(question.id);
      if (analytics) {
        const response = analytics.responses.find(r => r.respondentId === respondentId);
        return response?.answer || '/';
      }
      return '';
    }
    case 'closed_single':
    case 'closed_binary':
    case 'closed_multi': {
      const respondent = survey.respondents.find(r => r.id === respondentId);
      if (respondent) {
        const answer = respondent.originalData[question.rawHeader] ||
          respondent.originalData[question.cleanedHeader] || '';
        return answer || '/';
      }
      return '';
    }
    default:
      return '';
  }
}

/**
 * Convert column index to Excel column letter (1=A, 2=B, 27=AA, etc.)
 */
function colToLetter(col: number): string {
  let letter = '';
  let c = col;
  while (c > 0) {
    const mod = (c - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    c = Math.floor((c - 1) / 26);
  }
  return letter;
}

/**
 * Sheet 1: Export - Full transposed matrix (VALUES ONLY - staging layer)
 * Rows = questions, Columns = respondents
 * This is the ONLY place where raw data lives.
 */
function createExportSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  questionRowMap: Map<string, number>
): void {
  const sheet = workbook.addWorksheet('Export');
  const respondents = survey.respondents;

  // Row 1: Cognome header + surnames
  const row1Data = ['', 'Cognome'];
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    row1Data.push(surname);
  });
  const row1 = sheet.addRow(row1Data);
  styleHeaderRow(row1);

  // Row 2: Nome header + names
  const row2Data = ['', 'Nome'];
  respondents.forEach(r => {
    const nome = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.split(' ').slice(1).join(' ') || '';
    row2Data.push(nome);
  });
  const row2 = sheet.addRow(row2Data);
  styleHeaderRow(row2);

  // Track current row number for formula references
  let currentRow = 3;

  // Group questions by block
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Add block header row
    if (blockId !== null) {
      const pageRow = ['', `Page ${blockId}`];
      respondents.forEach(() => pageRow.push(''));
      const sectionRow = sheet.addRow(pageRow);
      styleSectionRow(sectionRow);
      currentRow++;
    }

    questions.forEach(question => {
      const rowData: (string | number)[] = [];

      // Column A: Question key
      rowData.push(question.questionKey || '');

      // Column B: Question text (cleaned)
      rowData.push(question.questionText);

      // Respondent values (raw data)
      respondents.forEach(r => {
        const value = getRespondentAnswer(survey, question, r.id);
        rowData.push(value);
      });

      const dataRow = sheet.addRow(rowData);
      styleDataRow(dataRow, question.type);

      // Map question key to row number for formula references
      if (question.questionKey) {
        questionRowMap.set(question.questionKey, currentRow);
      }
      // Also map by question ID for reliable lookup
      questionRowMap.set(question.id, currentRow);
      
      currentRow++;
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 8;
  sheet.getColumn(2).width = 80;
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 18;
  }

  sheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
}

/**
 * Sheet 2: Foglio2 - Same as Export but values only (no formulas)
 * This is a pure VALUES copy of the Export sheet
 */
function createFoglio2Sheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Foglio2');
  const respondents = survey.respondents;

  // Row 1: Cognome
  const row1Data = ['', 'Cognome'];
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    row1Data.push(surname);
  });
  const row1 = sheet.addRow(row1Data);
  styleHeaderRow(row1);

  // Row 2: Nome
  const row2Data = ['', 'Nome'];
  respondents.forEach(r => {
    const nome = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.split(' ').slice(1).join(' ') || '';
    row2Data.push(nome);
  });
  const row2 = sheet.addRow(row2Data);
  styleHeaderRow(row2);

  // All questions flat (values only)
  survey.questions.forEach(question => {
    const rowData: (string | number)[] = [
      question.questionKey || '',
      question.questionText
    ];

    respondents.forEach(r => {
      const value = getRespondentAnswer(survey, question, r.id);
      rowData.push(value);
    });

    sheet.addRow(rowData);
  });

  sheet.getColumn(1).width = 8;
  sheet.getColumn(2).width = 80;
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 18;
  }
}

/**
 * Sheet 3: Persone - Respondent metadata grid
 */
function createPersoneSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Persone');

  // Define metadata fields to look for
  const metadataFields = [
    'Cognome', 'Nome', 'CARICA', 'Carica',
    'COMITATO1', 'COMITATO2', 'COMITATO3', 'COMITATO4', 'COMITATO5', 'COMITATO6',
    'Comitato1', 'Comitato2', 'Comitato3', 'Comitato4', 'Comitato5', 'Comitato6',
    'RUOLO1', 'RUOLO2', 'RUOLO3', 'RUOLO4', 'RUOLO5', 'RUOLO6',
    'Ruolo1', 'Ruolo2', 'Ruolo3', 'Ruolo4', 'Ruolo5', 'Ruolo6',
    'TITOLO', 'Titolo', 'Email', 'email', 'ID Contact', 'ID Session'
  ];

  // Find which fields exist in the data
  const availableFields: string[] = [];
  const sampleRespondent = survey.respondents[0];
  if (sampleRespondent) {
    const dataKeys = Object.keys(sampleRespondent.originalData);
    metadataFields.forEach(field => {
      if (dataKeys.some(k => k.toLowerCase() === field.toLowerCase())) {
        const actualKey = dataKeys.find(k => k.toLowerCase() === field.toLowerCase());
        if (actualKey && !availableFields.includes(actualKey)) {
          availableFields.push(actualKey);
        }
      }
    });
  }

  // If no specific fields found, use Cognome, Nome at minimum
  if (availableFields.length === 0) {
    availableFields.push('Cognome', 'Nome');
  }

  // Header row
  const headers = ['#', ...availableFields];
  const headerRow = sheet.addRow(headers);
  styleHeaderRow(headerRow);

  // Data rows
  survey.respondents.forEach((respondent, idx) => {
    const rowData: (string | number)[] = [idx + 1];
    availableFields.forEach(field => {
      const value = respondent.originalData[field] || '';
      rowData.push(value);
    });
    sheet.addRow(rowData);
  });

  // Set column widths
  sheet.getColumn(1).width = 5;
  for (let i = 2; i <= availableFields.length + 1; i++) {
    sheet.getColumn(i).width = 20;
  }
}

/**
 * Sheet 4: estrazione per grafici  (with trailing space!)
 * FORMULA-DRIVEN - references Export sheet
 * 
 * Column A: =LEFT(B[row], SEARCH(" ", B[row]) - 1) or question key
 * Column B: question label
 * Column C: MEDIE formula (AVERAGEIF ignoring N/A)
 * Columns D..: VLOOKUP formulas to Export
 */
function createEstrazioneGraficiSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  questionRowMap: Map<string, number>
): void {
  const sheet = workbook.addWorksheet('estrazione per grafici ');
  const respondents = survey.respondents;

  // Headers - these are values, not formulas
  const headers = ['Chiave', 'Domanda', 'MEDIE'];
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    headers.push(surname);
  });
  const headerRow = sheet.addRow(headers);
  styleHeaderRow(headerRow);

  // Group and sort scale questions by block
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  // Current row for formula building
  let currentRow = 2;

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Add block header (value, not formula)
    const blockRowData: (string | ExcelJS.CellFormulaValue)[] = ['', getBlockDisplayName(blockId), ''];
    respondents.forEach(() => blockRowData.push(''));
    const sectionRow = sheet.addRow(blockRowData);
    styleSectionRow(sectionRow);
    currentRow++;

    // Sort questions within block by subId
    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const exportRow = questionRowMap.get(question.id);
      
      if (!exportRow) {
        // Fallback to values if no Export row found
        const analytics = survey.scaleAnalytics.get(question.id);
        const rowData: (string | number)[] = [
          question.questionKey || '',
          question.questionText,
          analytics?.mean || 0
        ];
        respondents.forEach(r => {
          if (analytics) {
            const value = analytics.respondentValues[r.id];
            rowData.push(value !== null ? value : 'N/A');
          } else {
            rowData.push('');
          }
        });
        sheet.addRow(rowData);
        currentRow++;
        return;
      }

      // Build formula-driven row
      const row = sheet.addRow([]);
      
      // Column A (1): Question key - formula or value
      if (question.questionKey) {
        row.getCell(1).value = question.questionKey;
      } else {
        // Formula to extract key from question text
        row.getCell(1).value = { 
          formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),"")` 
        };
      }

      // Column B (2): Question text (value from Export via formula)
      row.getCell(2).value = { formula: `Export!B${exportRow}` };

      // Column C (3): MEDIE formula - average of respondent values, ignoring N/A
      // Respondent data starts at column C (3) in Export, ends at column 2 + respondents.length
      const startCol = colToLetter(3);
      const endCol = colToLetter(2 + respondents.length);
      row.getCell(3).value = { 
        formula: `AVERAGEIF(Export!${startCol}${exportRow}:${endCol}${exportRow},"<>N/A")` 
      };
      row.getCell(3).font = { bold: true };
      row.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' }
      };
      row.getCell(3).alignment = { horizontal: 'center' };

      // Columns D.. (4+): VLOOKUP formulas to get respondent values from Export
      respondents.forEach((_, idx) => {
        const exportCol = colToLetter(3 + idx); // Column C onwards in Export
        const cellCol = 4 + idx;
        row.getCell(cellCol).value = { 
          formula: `IFERROR(Export!${exportCol}${exportRow},"")` 
        };
        row.getCell(cellCol).alignment = { horizontal: 'center' };
      });

      currentRow++;
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 8;
  sheet.getColumn(2).width = 60;
  sheet.getColumn(3).width = 10;
  for (let i = 4; i <= 3 + respondents.length; i++) {
    sheet.getColumn(i).width = 12;
  }

  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
}

/**
 * Sheet 5: per pdf  (with trailing space!)
 * FORMULA-DRIVEN - references estrazione per grafici sheet
 */
function createPerPdfSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  questionRowMap: Map<string, number>
): void {
  const sheet = workbook.addWorksheet('per pdf ');
  const respondents = survey.respondents;
  const estrazioneSheetName = "'estrazione per grafici '";

  // Title
  const titleRow = sheet.addRow(['Riepilogo Survey']);
  titleRow.font = { bold: true, size: 16 };
  titleRow.height = 30;

  // Summary stats (values - these are metadata)
  sheet.addRow(['']);
  sheet.addRow(['Rispondenti totali:', survey.respondents.length]);
  sheet.addRow(['Domande scala:', survey.questions.filter(q => q.type === 'scale_1_10_na').length]);
  sheet.addRow(['Domande aperte:', survey.questions.filter(q => q.type === 'open_text').length]);
  sheet.addRow(['']);

  // Scale questions section
  const scaleHeader = sheet.addRow(['DOMANDE SCALA 1-10']);
  scaleHeader.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  scaleHeader.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2563EB' }
  };

  // Headers for scale section
  const scaleHeaders = ['Chiave', 'Domanda', 'MEDIE'];
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    scaleHeaders.push(surname);
  });
  // Add count headers as explicit strings
  SCALE_ORDER.forEach(val => scaleHeaders.push(val));
  
  const scaleHeaderRow = sheet.addRow(scaleHeaders);
  styleHeaderRow(scaleHeaderRow);

  // Find corresponding rows in estrazione per grafici sheet
  // Since estrazione sheet mirrors scale questions, we track row mapping
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  // Track estrazione row (starts at 2 after header)
  let estrazioneRow = 2;
  let currentPdfRow = 9; // After title and headers

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Block header
    const blockRowData = ['', getBlockDisplayName(blockId)];
    for (let i = 2; i < scaleHeaders.length; i++) blockRowData.push('');
    const sectionRow = sheet.addRow(blockRowData);
    styleSectionRow(sectionRow);
    currentPdfRow++;
    estrazioneRow++; // Block header exists in estrazione too

    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const row = sheet.addRow([]);
      
      // Column A: Key (formula reference to estrazione)
      row.getCell(1).value = { formula: `${estrazioneSheetName}!A${estrazioneRow}` };
      
      // Column B: Question text (formula reference)
      row.getCell(2).value = { formula: `${estrazioneSheetName}!B${estrazioneRow}` };
      
      // Column C: MEDIE (formula reference)
      row.getCell(3).value = { formula: `${estrazioneSheetName}!C${estrazioneRow}` };
      row.getCell(3).font = { bold: true };
      row.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' }
      };
      row.getCell(3).alignment = { horizontal: 'center' };

      // Respondent values (formula references)
      respondents.forEach((_, idx) => {
        const estrazioneCol = colToLetter(4 + idx);
        row.getCell(4 + idx).value = { formula: `${estrazioneSheetName}!${estrazioneCol}${estrazioneRow}` };
        row.getCell(4 + idx).alignment = { horizontal: 'center' };
      });

      // Count columns (10..1, N/A) - COUNTIF formulas over estrazione respondent range
      const respStartCol = colToLetter(4);
      const respEndCol = colToLetter(3 + respondents.length);
      
      SCALE_ORDER.forEach((scaleVal, idx) => {
        const cellIdx = 4 + respondents.length + idx;
        const criteriaValue = scaleVal === 'N/A' ? '"N/A"' : scaleVal;
        row.getCell(cellIdx).value = { 
          formula: `COUNTIF(${estrazioneSheetName}!${respStartCol}${estrazioneRow}:${respEndCol}${estrazioneRow},${criteriaValue})` 
        };
        row.getCell(cellIdx).alignment = { horizontal: 'center' };
        row.getCell(cellIdx).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF3F4F6' }
        };
      });

      currentPdfRow++;
      estrazioneRow++;
    });
  });

  // Open questions section - values (not formula-driven, as they're text)
  sheet.addRow(['']);
  const openHeader = sheet.addRow(['DOMANDE APERTE']);
  openHeader.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  openHeader.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF059669' }
  };

  const openQuestions = survey.questions.filter(q => q.type === 'open_text');
  openQuestions.forEach(question => {
    const analytics = survey.openAnalytics.get(question.id);
    if (!analytics) return;

    // Question header
    const qRow = sheet.addRow([question.questionKey || '', question.questionText]);
    qRow.font = { bold: true };
    qRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0F2F1' }
    };

    // Responses (values - open text can't be formula-driven)
    analytics.responses.forEach(response => {
      const respRow = sheet.addRow(['', `${response.respondentName}: ${response.answer}`]);
      respRow.getCell(2).alignment = { wrapText: true };
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 8;
  sheet.getColumn(2).width = 60;
  sheet.getColumn(3).width = 10;
  for (let i = 4; i <= scaleHeaders.length; i++) {
    sheet.getColumn(i).width = 10;
  }
}

/**
 * Style helpers
 */
function styleHeaderRow(row: ExcelJS.Row): void {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2563EB' }
  };
  row.alignment = { vertical: 'middle', horizontal: 'center' };
  row.height = 28;
}

function styleSectionRow(row: ExcelJS.Row): void {
  row.font = { bold: true, italic: true };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E7FF' }
  };
  row.height = 24;
}

function styleDataRow(row: ExcelJS.Row, questionType: string): void {
  row.getCell(1).alignment = { horizontal: 'left', vertical: 'top' };
  row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

  if (questionType === 'scale_1_10_na') {
    for (let i = 3; i <= row.cellCount; i++) {
      row.getCell(i).alignment = { horizontal: 'center' };
    }
  }

  if (questionType === 'open_text') {
    for (let i = 3; i <= row.cellCount; i++) {
      row.getCell(i).alignment = { wrapText: true, vertical: 'top' };
    }
  }
}
