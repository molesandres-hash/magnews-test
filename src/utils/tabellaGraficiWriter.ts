import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

// Scale order - these MUST be written as strings to avoid "V10" display issue
const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

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
 * Generate tabella_grafici_scala10_NA_GENERATA.xlsx
 * Scale questions table with means and distribution counts
 * 
 * CRITICAL: Count column headers must be STRINGS, not numbers
 * to prevent Excel from displaying "V10" instead of "10"
 */
export async function generateTabellaGrafici(survey: ParsedSurvey): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  createMainSheet(workbook, survey);

  // Generate and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  saveAs(blob, 'tabella_grafici_scala10_NA_GENERATA.xlsx');
}

/**
 * Create the main sheet with scale questions data
 * Uses FORMULAS for count columns (COUNTIF)
 */
function createMainSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Foglio1');
  const respondents = survey.respondents;
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');

  // Calculate column positions
  const respondentStartCol = 3; // Column C
  const respondentEndCol = 2 + respondents.length;
  const countStartCol = respondentEndCol + 1;

  // Build headers: Domanda | MEDIE | [respondents...] | 10 | 9 | 8 | ... | 1 | N/A
  // CRITICAL: Headers must be proper values
  const headerRow = sheet.addRow([]);
  
  // Column A: Domanda label
  headerRow.getCell(1).value = 'Domanda';
  
  // Column B: MEDIE label
  headerRow.getCell(2).value = 'MEDIE';
  
  // Respondent columns (using surnames)
  respondents.forEach((r, idx) => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || r.displayName;
    headerRow.getCell(respondentStartCol + idx).value = surname;
  });
  
  // Count column headers - MUST be written as TEXT strings to prevent "V10" display
  SCALE_ORDER.forEach((scaleVal, idx) => {
    const cell = headerRow.getCell(countStartCol + idx);
    // Explicitly set as rich text or with text format to ensure string display
    cell.value = scaleVal;
    cell.numFmt = '@'; // Text format - prevents Excel auto-conversion
  });

  styleHeaderRow(headerRow);

  // Group questions by block
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  const totalCols = countStartCol + SCALE_ORDER.length - 1;

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Add block header
    const blockName = getBlockDisplayName(blockId);
    const blockRow = sheet.addRow([blockName]);
    styleBlockHeader(blockRow, totalCols);

    // Sort questions by subId within block
    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) return;

      const currentRowNum = sheet.rowCount + 1;
      const dataRow = sheet.addRow([]);

      // Column A: Question text
      dataRow.getCell(1).value = question.questionText;
      dataRow.getCell(1).alignment = { wrapText: true, vertical: 'top' };

      // Column B: MEDIE - formula to average respondent values (ignoring N/A)
      const respStartColLetter = colToLetter(respondentStartCol);
      const respEndColLetter = colToLetter(respondentEndCol);
      dataRow.getCell(2).value = { 
        formula: `AVERAGEIF(${respStartColLetter}${currentRowNum}:${respEndColLetter}${currentRowNum},"<>N/A")` 
      };
      dataRow.getCell(2).alignment = { horizontal: 'center' };
      dataRow.getCell(2).font = { bold: true };
      dataRow.getCell(2).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' }
      };

      // Respondent value columns - write actual values
      respondents.forEach((r, idx) => {
        const value = analytics.respondentValues[r.id];
        const cell = dataRow.getCell(respondentStartCol + idx);
        cell.value = value !== null ? value : 'N/A';
        cell.alignment = { horizontal: 'center' };
      });

      // Count columns - COUNTIF formulas over respondent range
      SCALE_ORDER.forEach((scaleVal, idx) => {
        const cell = dataRow.getCell(countStartCol + idx);
        const criteriaValue = scaleVal === 'N/A' ? '"N/A"' : scaleVal;
        cell.value = { 
          formula: `COUNTIF(${respStartColLetter}${currentRowNum}:${respEndColLetter}${currentRowNum},${criteriaValue})` 
        };
        cell.alignment = { horizontal: 'center' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF3F4F6' }
        };
      });

      // Add borders to all cells
      for (let i = 1; i <= totalCols; i++) {
        dataRow.getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };
      }
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 60; // Question text
  sheet.getColumn(2).width = 10; // MEDIE
  
  // Respondent columns
  for (let i = respondentStartCol; i <= respondentEndCol; i++) {
    sheet.getColumn(i).width = 14;
  }
  
  // Count columns (10, 9, 8, ..., 1, N/A)
  for (let i = countStartCol; i <= totalCols; i++) {
    sheet.getColumn(i).width = 6;
  }

  // Freeze header row
  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
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
  row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  row.height = 30;
}

function styleBlockHeader(row: ExcelJS.Row, colCount: number): void {
  row.font = { bold: true, size: 12 };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E7FF' }
  };
  row.height = 28;

  // Merge across all columns for block header
  const sheet = row.worksheet;
  sheet.mergeCells(row.number, 1, row.number, colCount);
}
