import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

/**
 * Generate and download the Excel report
 */
export async function generateExcelReport(survey: ParsedSurvey): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  // Sheet 1: Scale questions with counts
  createScaleSheet(workbook, survey, true);

  // Sheet 2: Scale questions without counts (for charts)
  createScaleSheet(workbook, survey, false);

  // Sheet 3: Open questions
  createOpenSheet(workbook, survey);

  // Sheet 4: Closed questions
  createClosedSheet(workbook, survey);

  // Sheet 5: Metadata
  createMetadataSheet(workbook, survey);

  // Generate and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
  saveAs(blob, `Report_${survey.metadata.fileName.replace('.csv', '')}.xlsx`);
}

/**
 * Create scale questions sheet
 */
function createScaleSheet(
  workbook: ExcelJS.Workbook,
  survey: ParsedSurvey,
  includeCounts: boolean
): void {
  const sheetName = includeCounts 
    ? 'tabella grafici scala 10 con NA' 
    : 'estrazione per grafici';
  const sheet = workbook.addWorksheet(sheetName);

  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const respondents = survey.respondents;

  // Build headers
  const headers = ['Domanda', 'MEDIE'];
  respondents.forEach(r => headers.push(r.displayName));
  if (includeCounts) {
    headers.push(...SCALE_ORDER);
  }

  // Add header row
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
    const blockRow = sheet.addRow([getBlockDisplayName(blockId)]);
    styleBlockHeader(blockRow, headers.length);

    // Add questions
    questions.forEach(question => {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) return;

      const rowData: (string | number)[] = [
        question.questionText,
        analytics.mean,
      ];

      // Add respondent values
      respondents.forEach(r => {
        const value = analytics.respondentValues[r.id];
        rowData.push(value !== null ? value : '');
      });

      // Add counts
      if (includeCounts) {
        SCALE_ORDER.forEach(key => {
          rowData.push(analytics.counts[key] || 0);
        });
      }

      const dataRow = sheet.addRow(rowData);
      styleDataRow(dataRow, headers.length);
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 60;
  sheet.getColumn(2).width = 10;
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 14;
  }
  if (includeCounts) {
    for (let i = 3 + respondents.length; i <= headers.length; i++) {
      sheet.getColumn(i).width = 6;
    }
  }

  // Freeze first row
  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
}

/**
 * Create open questions sheet
 */
function createOpenSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('aperte');

  const headers = ['Blocco', 'Domanda', 'Rispondente', 'Risposta'];
  const headerRow = sheet.addRow(headers);
  styleHeaderRow(headerRow);

  const openQuestions = survey.questions.filter(q => q.type === 'open_text');

  openQuestions.forEach(question => {
    const analytics = survey.openAnalytics.get(question.id);
    if (!analytics) return;

    analytics.responses.forEach(response => {
      const row = sheet.addRow([
        question.blockId !== null ? `Blocco ${question.blockId}` : 'N/D',
        question.questionText,
        response.respondentName,
        response.answer,
      ]);
      row.getCell(4).alignment = { wrapText: true, vertical: 'top' };
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 50;
  sheet.getColumn(3).width = 25;
  sheet.getColumn(4).width = 80;

  // Add auto filter
  sheet.autoFilter = { from: 'A1', to: 'D1' };
}

/**
 * Create closed questions sheet
 */
function createClosedSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('chiuse');

  const headers = ['Blocco', 'Domanda', 'Tipo', 'Opzione', 'Conteggio', 'Percentuale'];
  const headerRow = sheet.addRow(headers);
  styleHeaderRow(headerRow);

  const closedQuestions = survey.questions.filter(
    q => q.type === 'closed_single' || q.type === 'closed_binary' || q.type === 'closed_multi'
  );

  closedQuestions.forEach(question => {
    const analytics = survey.closedAnalytics.get(question.id);
    if (!analytics) return;

    analytics.options.forEach(option => {
      sheet.addRow([
        question.blockId !== null ? `Blocco ${question.blockId}` : 'N/D',
        question.questionText,
        question.type === 'closed_multi' ? 'Multipla' : 'Singola',
        option.option,
        option.count,
        `${option.percent}%`,
      ]);
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 50;
  sheet.getColumn(3).width = 12;
  sheet.getColumn(4).width = 30;
  sheet.getColumn(5).width = 12;
  sheet.getColumn(6).width = 12;

  // Add auto filter
  sheet.autoFilter = { from: 'A1', to: 'F1' };
}

/**
 * Create metadata sheet
 */
function createMetadataSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('metadata');

  const data = [
    ['File', survey.metadata.fileName],
    ['Data elaborazione', survey.metadata.parsedAt.toLocaleString('it-IT')],
    ['Righe totali', survey.metadata.totalRows],
    ['Risposte complete', survey.metadata.completedCount],
    ['Escluse', survey.metadata.excludedCount],
    ['Sessioni test', survey.metadata.testSessionCount],
    ['Domande totali', survey.questions.length],
    ['Domande scala', survey.questions.filter(q => q.type === 'scale_1_10_na').length],
    ['Domande aperte', survey.questions.filter(q => q.type === 'open_text').length],
    ['Domande chiuse', survey.questions.filter(
      q => q.type === 'closed_single' || q.type === 'closed_binary' || q.type === 'closed_multi'
    ).length],
    ['', ''],
    ['AVVISI', ''],
  ];

  survey.metadata.warnings.forEach(warning => {
    data.push(['', warning]);
  });

  data.forEach(row => sheet.addRow(row));

  sheet.getColumn(1).width = 25;
  sheet.getColumn(2).width = 60;

  // Style the header rows
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(12).font = { bold: true, color: { argb: 'FFCC6600' } };
}

/**
 * Style helper functions
 */
function styleHeaderRow(row: ExcelJS.Row): void {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2563EB' },
  };
  row.alignment = { vertical: 'middle', horizontal: 'center' };
  row.height = 30;
}

function styleBlockHeader(row: ExcelJS.Row, colCount: number): void {
  row.font = { bold: true, size: 12 };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E7FF' },
  };
  row.height = 28;
  
  // Merge cells for block header
  const sheet = row.worksheet;
  sheet.mergeCells(row.number, 1, row.number, colCount);
}

function styleDataRow(row: ExcelJS.Row, colCount: number): void {
  row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
  row.getCell(2).alignment = { horizontal: 'center' };
  
  // Add light borders
  for (let i = 1; i <= colCount; i++) {
    row.getCell(i).border = {
      top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
      bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    };
  }
}
