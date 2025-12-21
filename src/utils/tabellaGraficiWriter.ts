import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

/**
 * Generate tabella_grafici_scala10_NA_GENERATA.xlsx
 * Scale questions table with means and distribution counts
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
 */
function createMainSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Foglio1');
  const respondents = survey.respondents;
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');

  // Build headers: Domanda | MEDIE | [respondents...] | 10 | 9 | 8 | ... | 1 | N/A
  const headers: string[] = ['Domanda', 'MEDIE'];
  
  // Add respondent columns (using surnames for display)
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || r.displayName;
    headers.push(surname);
  });
  
  // Add count columns
  headers.push(...SCALE_ORDER);

  const headerRow = sheet.addRow(headers);
  styleHeaderRow(headerRow);

  // Group questions by block
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Add block header
    const blockName = getBlockDisplayName(blockId);
    const blockRowData: (string | number)[] = [blockName];
    // Fill remaining columns with empty strings
    for (let i = 1; i < headers.length; i++) {
      blockRowData.push('');
    }
    const blockRow = sheet.addRow(blockRowData);
    styleBlockHeader(blockRow, headers.length);

    // Sort questions by subId within block
    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) return;

      // Build row: Question text | Mean | [respondent values...] | [counts...]
      const rowData: (string | number)[] = [
        question.questionText,
        analytics.mean
      ];

      // Add respondent values
      respondents.forEach(r => {
        const value = analytics.respondentValues[r.id];
        rowData.push(value !== null ? value : 'N/A');
      });

      // Add counts for each scale value
      SCALE_ORDER.forEach(key => {
        rowData.push(analytics.counts[key] || 0);
      });

      const dataRow = sheet.addRow(rowData);
      styleDataRow(dataRow, headers.length, respondents.length);
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 60; // Question text
  sheet.getColumn(2).width = 10; // MEDIE
  
  // Respondent columns
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 14;
  }
  
  // Count columns (10, 9, 8, ..., 1, N/A)
  for (let i = 3 + respondents.length; i <= headers.length; i++) {
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

function styleDataRow(row: ExcelJS.Row, colCount: number, respondentCount: number): void {
  // Question text
  row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
  
  // MEDIE - highlighted
  row.getCell(2).alignment = { horizontal: 'center' };
  row.getCell(2).font = { bold: true };
  row.getCell(2).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFEF3C7' } // Amber light
  };

  // Respondent values - centered
  for (let i = 3; i <= 2 + respondentCount; i++) {
    row.getCell(i).alignment = { horizontal: 'center' };
  }

  // Count columns - centered with subtle background
  for (let i = 3 + respondentCount; i <= colCount; i++) {
    row.getCell(i).alignment = { horizontal: 'center' };
    row.getCell(i).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF3F4F6' }
    };
  }

  // Add borders
  for (let i = 1; i <= colCount; i++) {
    row.getCell(i).border = {
      top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
      bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
      left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
      right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
    };
  }
}
