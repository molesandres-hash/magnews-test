import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

/**
 * Generate file_per_questionari_GENERATO.xlsx
 * 
 * CRITICAL ARCHITECTURE:
 * - Export: ONLY sheet where raw data is written
 * - Foglio2: Values-only copy of Export
 * - Persone, estrazione per grafici, per pdf: FORMULA-DRIVEN (reference Export)
 * 
 * The formulas in downstream sheets use VLOOKUP to Export.
 * We must match the exact layout of the original OK file.
 */
export async function generateFilePerQuestionari(survey: ParsedSurvey): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  const respondents = survey.respondents;
  const numRespondents = respondents.length;
  
  // FIXED range to column T as per the OK file specification
  // This supports up to ~17 respondents
  const exportRange = 'Export!$A$1:$T$600';

  // Sheet 1: Export - THE ONLY SHEET WITH RAW DATA
  const exportRowMap = createExportSheet(workbook, survey);

  // Sheet 2: Foglio2 - Values copy of Export
  createFoglio2Sheet(workbook, survey);

  // Sheet 3: Persone - FORMULA-DRIVEN (VLOOKUP to Export)
  createPersoneSheet(workbook, survey, exportRange);

  // Sheet 4: estrazione per grafici  - FORMULA-DRIVEN
  createEstrazioneGraficiSheet(workbook, survey, exportRange, numRespondents);

  // Sheet 5: per pdf  - FORMULA-DRIVEN (references estrazione per grafici)
  createPerPdfSheet(workbook, survey, exportRange, numRespondents);

  // Generate and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  saveAs(blob, 'file_per_questionari_GENERATO.xlsx');
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
 * Sheet 1: Export - THE ONLY SHEET WITH RAW DATA
 * 
 * Layout (matches OK file):
 * - Column A: Key (can be formula or value matching LEFT(B,SEARCH(" ",B)-1))
 * - Column B: Label (Cognome, Nome, question text with number prefix)
 * - Columns C onwards: Respondent answers
 * 
 * Returns map of questionKey -> row number for reference
 */
function createExportSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey
): Map<string, number> {
  const sheet = workbook.addWorksheet('Export');
  const respondents = survey.respondents;
  const rowMap = new Map<string, number>();

  // Row 1: Cognome
  let currentRow = 1;
  const row1 = sheet.getRow(currentRow);
  row1.getCell(1).value = { formula: 'LEFT(B1,SEARCH(" ",B1)-1)' }; // "Cognome" key
  row1.getCell(2).value = 'Cognome';
  respondents.forEach((r, idx) => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    row1.getCell(3 + idx).value = surname;
  });
  styleHeaderRow(row1);
  row1.commit();

  // Row 2: Nome
  currentRow = 2;
  const row2 = sheet.getRow(currentRow);
  row2.getCell(1).value = { formula: 'LEFT(B2,SEARCH(" ",B2)-1)' }; // "Nome" key
  row2.getCell(2).value = 'Nome';
  respondents.forEach((r, idx) => {
    const nome = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.split(' ').slice(1).join(' ') || '';
    row2.getCell(3 + idx).value = nome;
  });
  styleHeaderRow(row2);
  row2.commit();

  currentRow = 3;

  // Group questions by block
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Add block/page header row
    if (blockId !== null) {
      const pageRow = sheet.getRow(currentRow);
      pageRow.getCell(1).value = { formula: `LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1)` };
      pageRow.getCell(2).value = `Page ${blockId}`;
      styleSectionRow(pageRow);
      pageRow.commit();
      currentRow++;
    }

    questions.forEach(question => {
      const row = sheet.getRow(currentRow);
      
      // Column A: Key formula (or direct value for reliability)
      // The formula =LEFT(B,SEARCH(" ",B)-1) extracts "4.1" from "4.1 Question text"
      const labelWithKey = question.questionKey 
        ? `${question.questionKey} ${question.questionText}`
        : question.questionText;
      
      row.getCell(1).value = { formula: `LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1)` };
      row.getCell(2).value = labelWithKey;

      // Columns C onwards: Respondent answers (raw values)
      respondents.forEach((r, idx) => {
        const value = getRespondentAnswer(survey, question, r.id);
        row.getCell(3 + idx).value = value;
      });

      row.commit();

      // Map for reference (use the key that will be computed by LEFT formula)
      if (question.questionKey) {
        rowMap.set(question.questionKey, currentRow);
      }
      rowMap.set(question.id, currentRow);
      
      currentRow++;
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 80;
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 18;
  }

  sheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
  
  return rowMap;
}

/**
 * Sheet 2: Foglio2 - VALUES-ONLY copy of Export
 * No formulas, just pure values
 */
function createFoglio2Sheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Foglio2');
  const respondents = survey.respondents;

  // Row 1: Cognome
  let currentRow = 1;
  const row1 = sheet.getRow(currentRow);
  row1.getCell(1).value = 'Cognome'; // Value, not formula
  row1.getCell(2).value = 'Cognome';
  respondents.forEach((r, idx) => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    row1.getCell(3 + idx).value = surname;
  });
  styleHeaderRow(row1);
  row1.commit();

  // Row 2: Nome
  currentRow = 2;
  const row2 = sheet.getRow(currentRow);
  row2.getCell(1).value = 'Nome';
  row2.getCell(2).value = 'Nome';
  respondents.forEach((r, idx) => {
    const nome = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.split(' ').slice(1).join(' ') || '';
    row2.getCell(3 + idx).value = nome;
  });
  styleHeaderRow(row2);
  row2.commit();

  currentRow = 3;

  // All questions (values only)
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    if (blockId !== null) {
      const pageRow = sheet.getRow(currentRow);
      pageRow.getCell(1).value = 'Page';
      pageRow.getCell(2).value = `Page ${blockId}`;
      styleSectionRow(pageRow);
      pageRow.commit();
      currentRow++;
    }

    questions.forEach(question => {
      const row = sheet.getRow(currentRow);
      
      const labelWithKey = question.questionKey 
        ? `${question.questionKey} ${question.questionText}`
        : question.questionText;
      
      row.getCell(1).value = question.questionKey || '';
      row.getCell(2).value = labelWithKey;

      respondents.forEach((r, idx) => {
        const value = getRespondentAnswer(survey, question, r.id);
        row.getCell(3 + idx).value = value;
      });

      row.commit();
      currentRow++;
    });
  });

  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 80;
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 18;
  }
}

/**
 * Sheet 3: Persone - FORMULA-DRIVEN with VLOOKUP
 * 
 * Layout (matches OK file):
 * - Row 21: Labels to search (Cognome, Nome, etc.) - one per column
 * - Rows 1-20: VLOOKUP formulas for each respondent (columns) x field (rows)
 * 
 * Formula pattern:
 * =IFERROR(VLOOKUP([col]$21,Export!$B:$M,[row_index],FALSE)," ")
 * 
 * Columns = respondents, Rows = metadata fields
 */
function createPersoneSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  exportRange: string
): void {
  const sheet = workbook.addWorksheet('Persone');
  const respondents = survey.respondents;
  const numRespondents = respondents.length;

  // Define the metadata fields to look up (these go in row 21 as search keys)
  // These are the labels in column B of Export that we want to find
  const fieldLabels = ['Cognome', 'Nome'];
  const numFields = fieldLabels.length;

  // Row 21: Headers/labels to search for (one per column A, B, C, ...)
  // Each column represents a respondent
  const row21 = sheet.getRow(21);
  for (let respIdx = 0; respIdx < numRespondents; respIdx++) {
    const col = respIdx + 1; // Column A=1, B=2, etc.
    // Each respondent column needs to know which field it's for
    // Actually, in the OK file: row 21 has the field names (Cognome, Nome, etc.)
    // and columns are respondents. Let me re-read the pattern.
    
    // From the OK file:
    // A21, B21, C21... contain the field labels (or respondent identifiers)
    // The VLOOKUP searches for these in Export
    
    // Actually the pattern is: columns = respondents, but the search key is always the same field
    // So we need to transpose: each COLUMN is a respondent, each ROW is a field
  }

  // Re-implementation based on OK file structure:
  // Columns A onwards = respondents
  // Rows 1-20 = data fields (Cognome, Nome, etc.)
  // Row 21 = search keys (the field labels)

  // Row 21: Field labels to search for
  for (let fieldIdx = 0; fieldIdx < fieldLabels.length; fieldIdx++) {
    row21.getCell(1 + fieldIdx).value = fieldLabels[fieldIdx];
  }
  row21.font = { italic: true };
  row21.commit();

  // For each respondent (column), create VLOOKUP formulas for each field (row)
  // Column A = respondent 1, Column B = respondent 2, etc.
  // The formula looks up the field label from row 21 in Export and returns the value from the respondent's column
  
  // BUT wait - in Export, the structure is:
  // Column A: Key, Column B: Label, Column C onwards: Respondent values
  // So to get a respondent's Cognome, we VLOOKUP "Cognome" in Export!$B:$M and return column 3+respIdx
  
  // Corrected structure for Persone:
  // Each ROW = a respondent
  // Each COLUMN = a field (Cognome, Nome, etc.)
  // This makes more sense for a "Persone" (people) listing
  
  // Headers in row 1
  const headerRow = sheet.getRow(1);
  headerRow.getCell(1).value = '#';
  fieldLabels.forEach((label, idx) => {
    headerRow.getCell(2 + idx).value = label;
  });
  styleHeaderRow(headerRow);
  headerRow.commit();

  // Row 21: Search keys (field labels) - used by VLOOKUP
  const searchKeyRow = sheet.getRow(21);
  fieldLabels.forEach((label, idx) => {
    searchKeyRow.getCell(2 + idx).value = label;
  });
  searchKeyRow.font = { italic: true, color: { argb: 'FF888888' } };
  searchKeyRow.commit();

  // Data rows: one per respondent with VLOOKUP formulas
  for (let respIdx = 0; respIdx < numRespondents; respIdx++) {
    const rowNum = 2 + respIdx; // Start from row 2
    const row = sheet.getRow(rowNum);
    
    row.getCell(1).value = respIdx + 1; // Index number

    // For each field, create a VLOOKUP formula
    // Formula: =IFERROR(VLOOKUP([field_cell]$21,Export!$B:$M,[colIndex],FALSE)," ")
    // Where colIndex = 2 for the first data column in Export (C), 3 for D, etc.
    // BUT we need to get the value for THIS specific respondent
    // In Export: respondent 1 is in column C (index 3), respondent 2 in D (index 4), etc.
    
    // The VLOOKUP searches for the field name (e.g., "Cognome") in Export column B
    // and returns the value from the respondent's column
    const exportColIndex = 2 + respIdx; // Column C=2 offset, D=3, etc. but VLOOKUP uses 1-based from range start

    for (let fieldIdx = 0; fieldIdx < fieldLabels.length; fieldIdx++) {
      const cellCol = 2 + fieldIdx;
      const colLetter = colToLetter(cellCol);
      
      // VLOOKUP formula: look up the field label from row 21 in Export!$B:$M
      // Return column index is 2 for respondent 1 (since range starts at B, C is column 2)
      // Actually: if range is $B:$M, column B=1, C=2, D=3, etc.
      // So for respondent 1 (column C in Export), return index = 2
      // For respondent 2 (column D in Export), return index = 3
      const returnColIndex = 1 + respIdx + 1; // +1 for 1-based, +respIdx for respondent offset
      
      row.getCell(cellCol).value = {
        formula: `IFERROR(VLOOKUP(${colLetter}$21,Export!$B:$M,${returnColIndex},FALSE)," ")`
      };
    }
    
    row.commit();
  }

  sheet.getColumn(1).width = 5;
  sheet.getColumn(2).width = 20;
  sheet.getColumn(3).width = 20;
}

/**
 * Sheet 4: estrazione per grafici  - FORMULA-DRIVEN
 * 
 * CRITICAL: This sheet uses VLOOKUP formulas to Export
 * - Column A: =LEFT(B{row},SEARCH(" ",B{row})-1)
 * - Column B: Question label (from template or written once)
 * - Column C: =IFERROR(SUMIF(D{row}:T{row},">0")/COUNTIF(D{row}:T{row},">0")," ")
 * - Columns D onwards: =IFERROR(VLOOKUP($A{row},Export!$A$1:$T$600,{colIndex},FALSE),"")
 */
function createEstrazioneGraficiSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  exportRange: string,
  numRespondents: number
): void {
  const sheet = workbook.addWorksheet('estrazione per grafici ');
  const respondents = survey.respondents;

  // Header row (values, not formulas)
  const headerRow = sheet.getRow(1);
  headerRow.getCell(1).value = 'Chiave';
  headerRow.getCell(2).value = 'Domanda';
  headerRow.getCell(3).value = 'MEDIE';
  respondents.forEach((r, idx) => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    headerRow.getCell(4 + idx).value = surname;
  });
  styleHeaderRow(headerRow);
  headerRow.commit();

  // FIXED column range for MEDIE formula (always D to T per OK file spec)
  const firstRespCol = 'D';
  const lastRespCol = 'T';

  let currentRow = 2;

  // Only scale questions for this sheet
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Block header
    const blockRow = sheet.getRow(currentRow);
    blockRow.getCell(1).value = '';
    blockRow.getCell(2).value = getBlockDisplayName(blockId);
    styleSectionRow(blockRow);
    blockRow.commit();
    currentRow++;

    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const row = sheet.getRow(currentRow);
      
      // Column A: Key formula
      row.getCell(1).value = { formula: `LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1)` };

      // Column B: Question label with key prefix
      const labelWithKey = question.questionKey 
        ? `${question.questionKey} ${question.questionText}`
        : question.questionText;
      row.getCell(2).value = labelWithKey;
      row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

      // Column C: MEDIE formula (average of values > 0, ignoring N/A)
      row.getCell(3).value = { 
        formula: `IFERROR(SUMIF(${firstRespCol}${currentRow}:${lastRespCol}${currentRow},">0")/COUNTIF(${firstRespCol}${currentRow}:${lastRespCol}${currentRow},">0")," ")` 
      };
      row.getCell(3).font = { bold: true };
      row.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' }
      };
      row.getCell(3).alignment = { horizontal: 'center' };

      // Columns D onwards: VLOOKUP formulas to Export
      // colIndex starts at 3 (column C in Export is first respondent)
      for (let i = 0; i < numRespondents; i++) {
        const colIndex = 3 + i; // Export column index for VLOOKUP
        row.getCell(4 + i).value = { 
          formula: `IFERROR(VLOOKUP($A${currentRow},${exportRange},${colIndex},FALSE),"")` 
        };
        row.getCell(4 + i).alignment = { horizontal: 'center' };
      }

      row.commit();
      currentRow++;
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 60;
  sheet.getColumn(3).width = 10;
  for (let i = 4; i <= 3 + numRespondents; i++) {
    sheet.getColumn(i).width = 12;
  }

  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
}

/**
 * Sheet 5: per pdf  - FORMULA-DRIVEN
 * 
 * This sheet shows one respondent at a time based on selector in H5
 * Uses IFS formula to pick the correct column from Export
 */
function createPerPdfSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  exportRange: string,
  numRespondents: number
): void {
  const sheet = workbook.addWorksheet('per pdf ');

  // Title area
  sheet.getRow(1).getCell(1).value = 'Riepilogo Survey';
  sheet.getRow(1).getCell(1).font = { bold: true, size: 16 };
  sheet.getRow(1).height = 30;

  // Respondent selector in H5
  sheet.getRow(5).getCell(8).value = 1; // Default to first respondent
  sheet.getRow(5).getCell(7).value = 'Rispondente:';
  
  // Add data validation for respondent selector (1 to numRespondents)
  sheet.getCell('H5').dataValidation = {
    type: 'whole',
    operator: 'between',
    formulae: [1, numRespondents],
    showErrorMessage: true,
    errorTitle: 'Valore non valido',
    error: `Inserire un numero da 1 a ${numRespondents}`
  };

  // Build the IFS formula for column selection
  // IFS($H$5=1,3,$H$5=2,4,$H$5=3,5,...)
  let ifsArgs = '';
  for (let i = 1; i <= numRespondents; i++) {
    const colIndex = 2 + i; // Column C=3 for respondent 1, D=4 for respondent 2, etc.
    ifsArgs += `$H$5=${i},${colIndex}`;
    if (i < numRespondents) ifsArgs += ',';
  }

  // Headers
  const headerRow = sheet.getRow(7);
  headerRow.getCell(1).value = 'Chiave';
  headerRow.getCell(2).value = 'Domanda';
  headerRow.getCell(3).value = 'MEDIE';
  headerRow.getCell(4).value = 'Risposta';
  styleHeaderRow(headerRow);

  let currentRow = 8;

  // Scale questions
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Block header
    const blockRow = sheet.getRow(currentRow);
    blockRow.getCell(1).value = '';
    blockRow.getCell(2).value = getBlockDisplayName(blockId);
    styleSectionRow(blockRow);
    currentRow++;

    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const row = sheet.getRow(currentRow);
      
      // Column A: Key (from column C which has the label)
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),"")` };

      // Column B: Question label
      const labelWithKey = question.questionKey 
        ? `${question.questionKey} ${question.questionText}`
        : question.questionText;
      row.getCell(2).value = labelWithKey;
      row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

      // Column C: MEDIE (reference to estrazione per grafici)
      row.getCell(3).value = { 
        formula: `IFERROR(VLOOKUP(A${currentRow},'estrazione per grafici '!$A:$C,3,FALSE),"")` 
      };
      row.getCell(3).font = { bold: true };
      row.getCell(3).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' }
      };
      row.getCell(3).alignment = { horizontal: 'center' };

      // Column D: Response for selected respondent (IFS formula)
      row.getCell(4).value = { 
        formula: `IFERROR(VLOOKUP(A${currentRow},${exportRange},IFS(${ifsArgs}),FALSE),"")` 
      };
      row.getCell(4).alignment = { horizontal: 'center' };

      currentRow++;
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 60;
  sheet.getColumn(3).width = 10;
  sheet.getColumn(4).width = 15;
  sheet.getColumn(7).width = 12;
  sheet.getColumn(8).width = 8;
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
