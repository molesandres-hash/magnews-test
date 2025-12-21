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
  row1.getCell(1).value = { formula: 'IFERROR(LEFT(B1,SEARCH(" ",B1)-1),B1)' }; // "Cognome" key
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
  row2.getCell(1).value = { formula: 'IFERROR(LEFT(B2,SEARCH(" ",B2)-1),B2)' }; // "Nome" key
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
      pageRow.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };
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
      
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };
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
 * Layout (matches OK file EXACTLY):
 * - COLUMNS A, B, C, D... = Respondents (one column per respondent)
 * - ROWS 1-20 = Data fields (VLOOKUP formulas pulling from Export)
 * - ROW 21 = Field labels to search (Cognome, Nome, Carica, Comitato 1, Ruolo 1, etc.)
 * 
 * Formula pattern for cell A1:
 * =IFERROR(VLOOKUP(A$21,Export!$B:$M,2,FALSE)," ")
 * 
 * The formula searches for the field label in Export column B,
 * and returns the value from the respondent's data column.
 */
function createPersoneSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  exportRange: string
): void {
  const sheet = workbook.addWorksheet('Persone');
  const respondents = survey.respondents;
  const numRespondents = respondents.length;

  // Define the metadata fields to look up (these are row labels in Export)
  // These match the structure of the original OK file
  const fieldLabels = [
    'Cognome',      // Row 1 in Persone
    'Nome',         // Row 2 in Persone
    'Carica',       // Row 3
    'Comitato 1',   // Row 4
    'Ruolo 1',      // Row 5
    'Comitato 2',   // Row 6
    'Ruolo 2',      // Row 7
    'Comitato 3',   // Row 8
    'Ruolo 3',      // Row 9
    'Comitato 4',   // Row 10
    'Ruolo 4',      // Row 11
    'Comitato 5',   // Row 12
    'Ruolo 5',      // Row 13
    'Genere',       // Row 14
    'Titolo',       // Row 15
    'Ruolo',        // Row 16
    'Email',        // Row 17
    'Telefono',     // Row 18
    'Indirizzo',    // Row 19
    'Note',         // Row 20
  ];
  const numFields = Math.min(fieldLabels.length, 20); // Max 20 rows of data

  // ROW 21: Field labels to search for (one per ROW, not column)
  // In the OK file structure, row 21 contains the search keys
  // Each column (A, B, C...) in row 21 has the same field labels
  // because each column represents a different respondent
  const row21 = sheet.getRow(21);
  for (let respIdx = 0; respIdx < numRespondents; respIdx++) {
    const colNum = respIdx + 1; // Column A=1, B=2, etc.
    // Row 21 for each column contains "the identifier" for that respondent
    // In the OK file, this is used as the lookup key
    row21.getCell(colNum).value = respondents[respIdx].displayName || `Rispondente ${respIdx + 1}`;
  }
  row21.font = { italic: true, color: { argb: 'FF888888' } };
  row21.commit();

  // ROWS 1-20: VLOOKUP formulas for each field
  // Each column represents a respondent, each row represents a field
  // Formula: =IFERROR(VLOOKUP([field_label],Export!$B:$M,[respColIndex],FALSE)," ")
  // 
  // In Export, the layout is:
  // - Column A: Key formula
  // - Column B: Label (Cognome, Nome, question text)
  // - Column C: Respondent 1 values
  // - Column D: Respondent 2 values
  // - etc.
  //
  // So for VLOOKUP: search in B column, return from column C/D/E... based on respondent

  for (let fieldIdx = 0; fieldIdx < numFields; fieldIdx++) {
    const rowNum = fieldIdx + 1; // Row 1, 2, 3...
    const row = sheet.getRow(rowNum);
    const fieldLabel = fieldLabels[fieldIdx];

    for (let respIdx = 0; respIdx < numRespondents; respIdx++) {
      const colNum = respIdx + 1; // Column A=1, B=2, C=3...
      const colLetter = colToLetter(colNum);
      
      // Return column index: for respondent 1, return col 2 (since range starts at B, C is index 2)
      // For respondent 2, return col 3 (D is index 3), etc.
      const returnColIndex = respIdx + 2;
      
      // Create the VLOOKUP formula
      // Searches for the field label (hardcoded) in Export!$B:$M and returns the respondent's value
      row.getCell(colNum).value = {
        formula: `IFERROR(VLOOKUP("${fieldLabel}",Export!$B:$T,${returnColIndex},FALSE)," ")`
      };
    }
    row.commit();
  }

  // Set column widths
  for (let i = 1; i <= numRespondents; i++) {
    sheet.getColumn(i).width = 18;
  }
  
  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 0 }];
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
  const sheet = workbook.addWorksheet('estrazione per grafici');
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

  // ALL questions for this sheet (not just scale, to include open-text)
  const allQuestions = survey.questions;
  const grouped = groupQuestionsByBlock(allQuestions);
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
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };

      // Column B: Question label with key prefix
      const labelWithKey = question.questionKey 
        ? `${question.questionKey} ${question.questionText}`
        : question.questionText;
      row.getCell(2).value = labelWithKey;
      row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

      // Column C: MEDIE formula (average of values > 0, ignoring N/A)
      // Only for scale questions; for open-text questions, leave empty or "-"
      if (question.type === 'scale_1_10_na') {
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
      } else {
        row.getCell(3).value = '-';
        row.getCell(3).font = { italic: true, color: { argb: 'FF888888' } };
        row.getCell(3).alignment = { horizontal: 'center' };
      }

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
 * 
 * Dashboard structure:
 * - H5: Respondent selector (1, 2, 3...)
 * - G8:H14: INDEX formulas showing respondent metadata from Persone sheet
 * - Column D: VLOOKUP+IFS formulas showing selected respondent's answers
 */
function createPerPdfSheet(
  workbook: ExcelJS.Workbook, 
  survey: ParsedSurvey,
  exportRange: string,
  numRespondents: number
): void {
  const sheet = workbook.addWorksheet('per pdf');

  // Title area
  sheet.getRow(1).getCell(1).value = 'Riepilogo Survey';
  sheet.getRow(1).getCell(1).font = { bold: true, size: 16 };
  sheet.getRow(1).height = 30;

  // Label for selector
  sheet.getRow(4).getCell(7).value = 'Seleziona partecipante:';
  sheet.getRow(4).getCell(7).font = { bold: true };

  // Respondent selector in H5
  sheet.getRow(5).getCell(8).value = 1; // Default to first respondent
  sheet.getRow(5).getCell(7).value = 'N°:';
  sheet.getCell('H5').font = { bold: true, size: 14 };
  sheet.getCell('H5').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFEF3C7' }
  };
  
  // Add data validation for respondent selector (1 to numRespondents)
  sheet.getCell('H5').dataValidation = {
    type: 'whole',
    operator: 'between',
    formulae: [1, numRespondents],
    showErrorMessage: true,
    errorTitle: 'Valore non valido',
    error: `Inserire un numero da 1 a ${numRespondents}`
  };

  // === G8:H14: INDEX formulas for respondent metadata from Persone sheet ===
  // The Persone sheet has: columns=respondents, rows=fields
  // Row 1=Cognome, Row 2=Nome, Row 3=Carica, etc.
  // To get data for respondent N, we use INDEX to pick the Nth column
  
  const metadataLabels = [
    { label: 'Cognome:', personeRow: 1 },
    { label: 'Nome:', personeRow: 2 },
    { label: 'Carica:', personeRow: 3 },
    { label: 'Comitato 1:', personeRow: 4 },
    { label: 'Ruolo 1:', personeRow: 5 },
    { label: 'Comitato 2:', personeRow: 6 },
    { label: 'Ruolo 2:', personeRow: 7 },
  ];

  metadataLabels.forEach((meta, idx) => {
    const rowNum = 8 + idx; // Rows 8, 9, 10, 11, 12, 13, 14
    const row = sheet.getRow(rowNum);
    
    // Column F: Label
    row.getCell(6).value = meta.label;
    row.getCell(6).font = { bold: true };
    row.getCell(6).alignment = { horizontal: 'right' };
    
    // Column G: INDEX formula to get value from Persone sheet
    // INDEX(Persone!1:1, H5) gets the value from row 1, column H5
    row.getCell(7).value = {
      formula: `INDEX(Persone!${meta.personeRow}:${meta.personeRow},'per pdf'!$H$5)`
    };
    row.getCell(7).alignment = { horizontal: 'left' };
  });

  // Build the IFS formula for column selection
  // IFS($H$5=1,3,$H$5=2,4,$H$5=3,5,...)
  let ifsArgs = '';
  for (let i = 1; i <= Math.min(numRespondents, 17); i++) {
    const colIndex = 2 + i; // Column C=3 for respondent 1, D=4 for respondent 2, etc.
    ifsArgs += `$H$5=${i},${colIndex}`;
    if (i < Math.min(numRespondents, 17)) ifsArgs += ',';
  }

  // Headers
  const headerRow = sheet.getRow(16);
  headerRow.getCell(1).value = 'Chiave';
  headerRow.getCell(2).value = 'Domanda';
  headerRow.getCell(3).value = 'MEDIE';
  headerRow.getCell(4).value = 'Risposta';
  styleHeaderRow(headerRow);

  let currentRow = 17;

  // ALL questions (not just scale, to include open-text responses)
  const allQuestions = survey.questions;
  const grouped = groupQuestionsByBlock(allQuestions);
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
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };

      // Column B: Question label
      const labelWithKey = question.questionKey 
        ? `${question.questionKey} ${question.questionText}`
        : question.questionText;
      row.getCell(2).value = labelWithKey;
      row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

      // Column C: MEDIE (reference to estrazione per grafici) - only for scale questions
      if (question.type === 'scale_1_10_na') {
        row.getCell(3).value = { 
          formula: `IFERROR(VLOOKUP(A${currentRow},'estrazione per grafici'!$A:$C,3,FALSE),"")` 
        };
        row.getCell(3).font = { bold: true };
        row.getCell(3).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFEF3C7' }
        };
        row.getCell(3).alignment = { horizontal: 'center' };
      } else {
        row.getCell(3).value = '-';
        row.getCell(3).font = { italic: true, color: { argb: 'FF888888' } };
        row.getCell(3).alignment = { horizontal: 'center' };
      }

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
  sheet.getColumn(4).width = 20;
  sheet.getColumn(6).width = 12;
  sheet.getColumn(7).width = 25;
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
