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

  // Sheet 1: Main data sheet - Questions as rows, respondents as columns
  createMainDataSheet(workbook, survey);

  // Sheet 2: Scale questions summary
  createScaleSummarySheet(workbook, survey);

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
 * Create main data sheet with questions as rows and respondents as columns
 * Format matches the reference: Cognome, Nome rows, then questions
 */
function createMainDataSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Dati Completi');
  const respondents = survey.respondents;

  // Row 1: Header labels + Respondent surnames (displayName is typically surname initial)
  const headerRow1 = ['Cognome', 'Cognome'];
  respondents.forEach(r => {
    // Try to extract surname from displayName or originalData
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName;
    headerRow1.push(surname);
  });
  const row1 = sheet.addRow(headerRow1);
  styleHeaderRow(row1);

  // Row 2: Nome labels + Respondent names/initials
  const headerRow2 = ['Nome', 'Nome'];
  respondents.forEach(r => {
    const nome = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.charAt(0);
    headerRow2.push(nome);
  });
  const row2 = sheet.addRow(headerRow2);
  styleHeaderRow(row2);

  // Group questions by block for organization
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  let currentSection = '';

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];
    
    questions.forEach(question => {
      const rowData: (string | number)[] = [];

      // Column A: Question key (e.g., "10.", "4.1")
      const questionKey = question.questionKey || question.id;
      rowData.push(questionKey);

      // Column B: Full question text
      rowData.push(question.questionText);

      // Respondent values
      respondents.forEach(r => {
        const value = getRespondentAnswer(survey, question, r.id);
        rowData.push(value);
      });

      // Add section header row if block changed
      const sectionName = getBlockDisplayName(blockId);
      if (sectionName !== currentSection && blockId !== null) {
        // Add a "Page" / section indicator row
        const pageRow = ['Page', `Page ${sectionName}`];
        respondents.forEach(() => pageRow.push(''));
        const sectionRow = sheet.addRow(pageRow);
        styleSectionRow(sectionRow, 2 + respondents.length);
        currentSection = sectionName;
      }

      const dataRow = sheet.addRow(rowData);
      styleDataRowMain(dataRow, question.type);
    });
  });

  // Set column widths
  sheet.getColumn(1).width = 8;  // Question key
  sheet.getColumn(2).width = 80; // Question text
  for (let i = 3; i <= 2 + respondents.length; i++) {
    sheet.getColumn(i).width = 18;
  }

  // Freeze first 2 rows and first 2 columns
  sheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
}

/**
 * Get respondent's answer for a question
 */
function getRespondentAnswer(survey: ParsedSurvey, question: QuestionInfo, respondentId: string): string | number {
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
      // For closed questions, find the respondent's answer from raw data
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
 * Create scale questions summary sheet
 */
function createScaleSummarySheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Riepilogo Scale');
  const respondents = survey.respondents;
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');

  // Headers
  const headers = ['Domanda', 'MEDIE'];
  respondents.forEach(r => headers.push(r.displayName));
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
    const blockRow = sheet.addRow([getBlockDisplayName(blockId)]);
    styleBlockHeader(blockRow, headers.length);

    questions.forEach(question => {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) return;

      const rowData: (string | number)[] = [
        question.questionText,
        analytics.mean,
      ];

      respondents.forEach(r => {
        const value = analytics.respondentValues[r.id];
        rowData.push(value !== null ? value : 'N/A');
      });

      SCALE_ORDER.forEach(key => {
        rowData.push(analytics.counts[key] || 0);
      });

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
  for (let i = 3 + respondents.length; i <= headers.length; i++) {
    sheet.getColumn(i).width = 6;
  }

  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
}

/**
 * Create open questions sheet
 */
function createOpenSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Aperte');

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

  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 50;
  sheet.getColumn(3).width = 25;
  sheet.getColumn(4).width = 80;

  sheet.autoFilter = { from: 'A1', to: 'D1' };
}

/**
 * Create closed questions sheet
 */
function createClosedSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Chiuse');

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

  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 50;
  sheet.getColumn(3).width = 12;
  sheet.getColumn(4).width = 30;
  sheet.getColumn(5).width = 12;
  sheet.getColumn(6).width = 12;

  sheet.autoFilter = { from: 'A1', to: 'F1' };
}

/**
 * Create metadata sheet
 */
function createMetadataSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  const sheet = workbook.addWorksheet('Metadata');

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
  
  const sheet = row.worksheet;
  sheet.mergeCells(row.number, 1, row.number, colCount);
}

function styleSectionRow(row: ExcelJS.Row, colCount: number): void {
  row.font = { bold: true, italic: true };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF3F4F6' },
  };
  row.height = 24;
}

function styleDataRow(row: ExcelJS.Row, colCount: number): void {
  row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
  row.getCell(2).alignment = { horizontal: 'center' };
  
  for (let i = 1; i <= colCount; i++) {
    row.getCell(i).border = {
      top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
      bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    };
  }
}

function styleDataRowMain(row: ExcelJS.Row, questionType: string): void {
  row.getCell(1).alignment = { horizontal: 'left', vertical: 'top' };
  row.getCell(2).alignment = { wrapText: true, vertical: 'top' };
  
  // Light blue for scale questions
  if (questionType === 'scale_1_10_na') {
    for (let i = 3; i <= row.cellCount; i++) {
      row.getCell(i).alignment = { horizontal: 'center' };
    }
  }
  
  // Wrap text for open questions
  if (questionType === 'open_text') {
    for (let i = 3; i <= row.cellCount; i++) {
      row.getCell(i).alignment = { wrapText: true, vertical: 'top' };
    }
  }
}
