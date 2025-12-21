import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

// Scale order - MUST be written as TEXT strings to prevent "V10" display
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
 * 
 * This is a SEPARATE workbook with scale question analysis.
 * 
 * CRITICAL: 
 * - Count column headers (10, 9, 8, ..., 1, N/A) must be TEXT to prevent "V10"
 * - Count values use COUNTIF formulas
 * - MEDIE uses SUMIF/COUNTIF formula
 */
export async function generateTabellaGrafici(survey: ParsedSurvey): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  createMainSheet(workbook, survey);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  saveAs(blob, 'tabella_grafici_scala10_NA_GENERATA.xlsx');
}

/**
 * Create the main sheet with scale questions data
 */
function createMainSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Foglio1');
  const respondents = survey.respondents;
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const numRespondents = respondents.length;

  // Column positions
  const respondentStartCol = 3; // Column C
  const respondentEndCol = 2 + numRespondents;
  const countStartCol = respondentEndCol + 1;
  const totalCols = countStartCol + SCALE_ORDER.length - 1;

  // Column letters for formulas
  const firstRespColLetter = colToLetter(respondentStartCol);
  const lastRespColLetter = colToLetter(respondentEndCol);

  // Header row
  const headerRow = sheet.getRow(1);
  
  // Column A: Domanda
  headerRow.getCell(1).value = 'Domanda';
  
  // Column B: MEDIE
  headerRow.getCell(2).value = 'MEDIE';
  
  // Respondent columns (surnames)
  respondents.forEach((r, idx) => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || r.displayName;
    headerRow.getCell(respondentStartCol + idx).value = surname;
  });
  
  // Count column headers - EXPLICIT TEXT FORMAT to prevent "V10"
  SCALE_ORDER.forEach((scaleVal, idx) => {
    const cell = headerRow.getCell(countStartCol + idx);
    cell.value = scaleVal;
    cell.numFmt = '@'; // Force text format
  });

  styleHeaderRow(headerRow);
  headerRow.commit();

  // Group questions by block
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  let currentRow = 2;

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Block header
    const blockName = getBlockDisplayName(blockId);
    const blockRow = sheet.getRow(currentRow);
    blockRow.getCell(1).value = blockName;
    styleBlockHeader(blockRow, totalCols);
    blockRow.commit();
    currentRow++;

    // Sort questions by subId
    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) return;

      const dataRow = sheet.getRow(currentRow);

      // Column A: Question text
      dataRow.getCell(1).value = question.questionText;
      dataRow.getCell(1).alignment = { wrapText: true, vertical: 'top' };

      // Column B: MEDIE formula (average of numeric values > 0)
      dataRow.getCell(2).value = { 
        formula: `IFERROR(SUMIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},">0")/COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},">0")," ")` 
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

      // Count columns - COUNTIF formulas
      SCALE_ORDER.forEach((scaleVal, idx) => {
        const cell = dataRow.getCell(countStartCol + idx);
        
        // For N/A, we count text "N/A"
        // For numbers, we count the numeric value
        if (scaleVal === 'N/A') {
          cell.value = { 
            formula: `COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},"N/A")` 
          };
        } else {
          cell.value = { 
            formula: `COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},${scaleVal})` 
          };
        }
        
        cell.alignment = { horizontal: 'center' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF3F4F6' }
        };
      });

      // Borders
      for (let i = 1; i <= totalCols; i++) {
        dataRow.getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };
      }

      dataRow.commit();
      currentRow++;
    });
  });

  // Column widths
  sheet.getColumn(1).width = 60;
  sheet.getColumn(2).width = 10;
  
  for (let i = respondentStartCol; i <= respondentEndCol; i++) {
    sheet.getColumn(i).width = 14;
  }
  
  for (let i = countStartCol; i <= totalCols; i++) {
    sheet.getColumn(i).width = 6;
  }

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

  const sheet = row.worksheet;
  sheet.mergeCells(row.number, 1, row.number, colCount);
}
