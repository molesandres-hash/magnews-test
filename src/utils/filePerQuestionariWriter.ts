import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

/**
 * Generate file_per_questionari_GENERATO.xlsx
 * Contains 5 sheets: Export, Foglio2, Persone, estrazione per grafici , per pdf 
 */
export async function generateFilePerQuestionari(survey: ParsedSurvey): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  // Sheet 1: Export - Full transposed matrix
  createExportSheet(workbook, survey);

  // Sheet 2: Foglio2 - Values-only copy
  createFoglio2Sheet(workbook, survey);

  // Sheet 3: Persone - Respondent metadata
  createPersoneSheet(workbook, survey);

  // Sheet 4: estrazione per grafici  (with trailing space)
  createEstrazioneGraficiSheet(workbook, survey);

  // Sheet 5: per pdf  (with trailing space)
  createPerPdfSheet(workbook, survey);

  // Generate and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  saveAs(blob, 'file_per_questionari_GENERATO.xlsx');
}

/**
 * Get respondent's answer for a question
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
 * Sheet 1: Export - Full transposed matrix
 * Rows = questions, Columns = respondents
 */
function createExportSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
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

  // Group questions by block
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  let currentSection = '';

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];
    const sectionName = getBlockDisplayName(blockId);

    // Add block header row
    if (sectionName !== currentSection && blockId !== null) {
      const pageRow = ['', `Page ${blockId}`];
      respondents.forEach(() => pageRow.push(''));
      const sectionRow = sheet.addRow(pageRow);
      styleSectionRow(sectionRow, 2 + respondents.length);
      currentSection = sectionName;
    }

    questions.forEach(question => {
      const rowData: (string | number)[] = [];

      // Column A: Question key
      rowData.push(question.questionKey || '');

      // Column B: Question text (cleaned)
      rowData.push(question.questionText);

      // Respondent values
      respondents.forEach(r => {
        const value = getRespondentAnswer(survey, question, r.id);
        rowData.push(value);
      });

      const dataRow = sheet.addRow(rowData);
      styleDataRow(dataRow, question.type);
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
 * Ordered extraction with MEDIE column
 */
function createEstrazioneGraficiSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  // Note: sheet name has trailing space as per requirement
  const sheet = workbook.addWorksheet('estrazione per grafici ');
  const respondents = survey.respondents;

  // Headers
  const headers = ['Chiave', 'Domanda', 'MEDIE'];
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    headers.push(surname);
  });
  const headerRow = sheet.addRow(headers);
  styleHeaderRow(headerRow);

  // Group and sort questions by block
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];

    // Add block header
    const blockRow = ['', getBlockDisplayName(blockId), ''];
    respondents.forEach(() => blockRow.push(''));
    const sectionRow = sheet.addRow(blockRow);
    styleSectionRow(sectionRow, 3 + respondents.length);

    // Sort questions within block by subId
    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

    sortedQuestions.forEach(question => {
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

      const dataRow = sheet.addRow(rowData);
      styleScaleDataRow(dataRow, 3 + respondents.length);
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
 * Presentation-ready sheet
 */
function createPerPdfSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey): void {
  // Note: sheet name has trailing space as per requirement
  const sheet = workbook.addWorksheet('per pdf ');
  const respondents = survey.respondents;

  // Title
  const titleRow = sheet.addRow(['Riepilogo Survey']);
  titleRow.font = { bold: true, size: 16 };
  titleRow.height = 30;

  // Summary stats
  sheet.addRow(['']);
  sheet.addRow(['Rispondenti totali:', survey.respondents.length]);
  sheet.addRow(['Domande scala:', survey.questions.filter(q => q.type === 'scale_1_10_na').length]);
  sheet.addRow(['Domande aperte:', survey.questions.filter(q => q.type === 'open_text').length]);
  sheet.addRow(['']);

  // Scale questions section
  const scaleHeader = sheet.addRow(['DOMANDE SCALA 1-10']);
  scaleHeader.font = { bold: true, size: 14 };
  scaleHeader.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2563EB' }
  };
  scaleHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };

  // Headers for scale section
  const scaleHeaders = ['Chiave', 'Domanda', 'MEDIE'];
  respondents.forEach(r => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
    scaleHeaders.push(surname);
  });
  scaleHeaders.push(...SCALE_ORDER);
  
  const scaleHeaderRow = sheet.addRow(scaleHeaders);
  styleHeaderRow(scaleHeaderRow);

  // Scale questions data
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
    const blockRow = ['', getBlockDisplayName(blockId)];
    const restCols = scaleHeaders.length - 2;
    for (let i = 0; i < restCols; i++) blockRow.push('');
    const sectionRow = sheet.addRow(blockRow);
    styleSectionRow(sectionRow, scaleHeaders.length);

    questions.forEach(question => {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) return;

      const rowData: (string | number)[] = [
        question.questionKey || '',
        question.questionText,
        analytics.mean
      ];

      respondents.forEach(r => {
        const value = analytics.respondentValues[r.id];
        rowData.push(value !== null ? value : 'N/A');
      });

      SCALE_ORDER.forEach(key => {
        rowData.push(analytics.counts[key] || 0);
      });

      const dataRow = sheet.addRow(rowData);
      styleScaleDataRow(dataRow, scaleHeaders.length);
    });
  });

  // Open questions section
  sheet.addRow(['']);
  const openHeader = sheet.addRow(['DOMANDE APERTE']);
  openHeader.font = { bold: true, size: 14 };
  openHeader.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF059669' }
  };
  openHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };

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

    // Responses
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

function styleSectionRow(row: ExcelJS.Row, colCount: number): void {
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

function styleScaleDataRow(row: ExcelJS.Row, colCount: number): void {
  row.getCell(1).alignment = { horizontal: 'left' };
  row.getCell(2).alignment = { wrapText: true, vertical: 'top' };
  row.getCell(3).alignment = { horizontal: 'center' };
  row.getCell(3).font = { bold: true };
  row.getCell(3).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFEF3C7' }
  };

  for (let i = 4; i <= colCount; i++) {
    row.getCell(i).alignment = { horizontal: 'center' };
  }
}
