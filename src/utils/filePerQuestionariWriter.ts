import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';
import { useTemplateStore } from '@/store/templateStore';
import { hexToArgb } from './templateColors';
import { generateBlockMeanChartPNG } from './excelChartHelper';

export async function generateFilePerQuestionari(survey: ParsedSurvey): Promise<void> {
  const template = useTemplateStore.getState().getActiveTemplate();
  const fontName = template?.fontFamily || 'Calibri';
  const headerArgb = template ? hexToArgb(template.primaryColor) : 'FF2563EB';
  const sectionArgb = template ? hexToArgb(template.secondaryColor) : 'FFE0E7FF';

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  const respondents = survey.respondents;
  const numRespondents = respondents.length;
  const exportRange = 'Export!$A$1:$T$600';

  const exportRowMap = createExportSheet(workbook, survey, fontName, headerArgb, sectionArgb, template?.logoBase64);
  createFoglio2Sheet(workbook, survey, fontName, headerArgb, sectionArgb);
  createPersoneSheet(workbook, survey, exportRange);
  createEstrazioneGraficiSheet(workbook, survey, exportRange, numRespondents, fontName, headerArgb, sectionArgb);
  createPerPdfSheet(workbook, survey, exportRange, numRespondents, fontName, headerArgb, sectionArgb);

  // Create "Grafici" sheet with embedded charts
  await createGraficiSheet(workbook, survey, template);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'file_per_questionari_GENERATO.xlsx');
}

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

function getRespondentAnswer(survey: ParsedSurvey, question: QuestionInfo, respondentId: string): string | number {
  switch (question.type) {
    case 'scale_1_10_na': {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (analytics) { const value = analytics.respondentValues[respondentId]; return value !== null ? value : 'N/A'; }
      return '';
    }
    case 'open_text': {
      const analytics = survey.openAnalytics.get(question.id);
      if (analytics) { const response = analytics.responses.find(r => r.respondentId === respondentId); return response?.answer || '/'; }
      return '';
    }
    case 'closed_single': case 'closed_binary': case 'closed_multi': {
      const respondent = survey.respondents.find(r => r.id === respondentId);
      if (respondent) { return respondent.originalData[question.rawHeader] || respondent.originalData[question.cleanedHeader] || '/'; }
      return '';
    }
    default: return '';
  }
}

function createExportSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey, fontName: string, headerArgb: string, sectionArgb: string, logoBase64?: string): Map<string, number> {
  const sheet = workbook.addWorksheet('Export');
  const respondents = survey.respondents;
  const rowMap = new Map<string, number>();

  // Logo
  if (logoBase64) {
    const logoData = logoBase64.replace(/^data:image\/(png|jpeg|jpg);base64,/, '');
    const ext = logoBase64.includes('image/png') ? 'png' : 'jpeg';
    const lastCol = 2 + respondents.length;
    const imageId = workbook.addImage({ base64: logoData, extension: ext as 'png' | 'jpeg' });
    sheet.addImage(imageId, { tl: { col: lastCol - 1, row: 0 }, ext: { width: 120, height: 60 } });
  }

  let currentRow = 1;
  const row1 = sheet.getRow(currentRow);
  row1.getCell(1).value = { formula: 'IFERROR(LEFT(B1,SEARCH(" ",B1)-1),B1)' };
  row1.getCell(2).value = 'Cognome';
  respondents.forEach((r, idx) => {
    row1.getCell(3 + idx).value = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || '';
  });
  styleHeaderRow(row1, fontName, headerArgb);
  row1.commit();

  currentRow = 2;
  const row2 = sheet.getRow(currentRow);
  row2.getCell(1).value = { formula: 'IFERROR(LEFT(B2,SEARCH(" ",B2)-1),B2)' };
  row2.getCell(2).value = 'Nome';
  respondents.forEach((r, idx) => {
    row2.getCell(3 + idx).value = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.split(' ').slice(1).join(' ') || '';
  });
  styleHeaderRow(row2, fontName, headerArgb);
  row2.commit();

  currentRow = 3;
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => { if (a === null) return 1; if (b === null) return -1; return a - b; });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];
    if (blockId !== null) {
      const pageRow = sheet.getRow(currentRow);
      pageRow.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };
      pageRow.getCell(2).value = `Page ${blockId}`;
      styleSectionRow(pageRow, fontName, sectionArgb);
      pageRow.commit();
      currentRow++;
    }
    questions.forEach(question => {
      const row = sheet.getRow(currentRow);
      const labelWithKey = question.questionKey ? `${question.questionKey} ${question.questionText}` : question.questionText;
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };
      row.getCell(2).value = labelWithKey;
      respondents.forEach((r, idx) => { row.getCell(3 + idx).value = getRespondentAnswer(survey, question, r.id); });
      row.commit();
      if (question.questionKey) rowMap.set(question.questionKey, currentRow);
      rowMap.set(question.id, currentRow);
      currentRow++;
    });
  });

  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 80;
  for (let i = 3; i <= 2 + respondents.length; i++) sheet.getColumn(i).width = 18;
  sheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
  return rowMap;
}

function createFoglio2Sheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey, fontName: string, headerArgb: string, sectionArgb: string): void {
  const sheet = workbook.addWorksheet('Foglio2');
  const respondents = survey.respondents;

  let currentRow = 1;
  const row1 = sheet.getRow(currentRow);
  row1.getCell(1).value = 'Cognome'; row1.getCell(2).value = 'Cognome';
  respondents.forEach((r, idx) => { row1.getCell(3 + idx).value = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || ''; });
  styleHeaderRow(row1, fontName, headerArgb); row1.commit();

  currentRow = 2;
  const row2 = sheet.getRow(currentRow);
  row2.getCell(1).value = 'Nome'; row2.getCell(2).value = 'Nome';
  respondents.forEach((r, idx) => { row2.getCell(3 + idx).value = r.originalData['Nome'] || r.originalData['nome'] || r.displayName.split(' ').slice(1).join(' ') || ''; });
  styleHeaderRow(row2, fontName, headerArgb); row2.commit();

  currentRow = 3;
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => { if (a === null) return 1; if (b === null) return -1; return a - b; });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];
    if (blockId !== null) {
      const pageRow = sheet.getRow(currentRow);
      pageRow.getCell(1).value = 'Page'; pageRow.getCell(2).value = `Page ${blockId}`;
      styleSectionRow(pageRow, fontName, sectionArgb); pageRow.commit(); currentRow++;
    }
    questions.forEach(question => {
      const row = sheet.getRow(currentRow);
      const labelWithKey = question.questionKey ? `${question.questionKey} ${question.questionText}` : question.questionText;
      row.getCell(1).value = question.questionKey || '';
      row.getCell(2).value = labelWithKey;
      respondents.forEach((r, idx) => { row.getCell(3 + idx).value = getRespondentAnswer(survey, question, r.id); });
      row.commit(); currentRow++;
    });
  });

  sheet.getColumn(1).width = 10; sheet.getColumn(2).width = 80;
  for (let i = 3; i <= 2 + respondents.length; i++) sheet.getColumn(i).width = 18;
}

function createPersoneSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey, exportRange: string): void {
  const sheet = workbook.addWorksheet('Persone');
  const respondents = survey.respondents;
  const numRespondents = respondents.length;
  const fieldLabels = ['Cognome','Nome','Carica','Comitato 1','Ruolo 1','Comitato 2','Ruolo 2','Comitato 3','Ruolo 3','Comitato 4','Ruolo 4','Comitato 5','Ruolo 5','Genere','Titolo','Ruolo','Email','Telefono','Indirizzo','Note'];
  const numFields = Math.min(fieldLabels.length, 20);

  const row21 = sheet.getRow(21);
  for (let respIdx = 0; respIdx < numRespondents; respIdx++) {
    row21.getCell(respIdx + 1).value = respondents[respIdx].displayName || `Rispondente ${respIdx + 1}`;
  }
  row21.font = { italic: true, color: { argb: 'FF888888' } };
  row21.commit();

  for (let fieldIdx = 0; fieldIdx < numFields; fieldIdx++) {
    const row = sheet.getRow(fieldIdx + 1);
    for (let respIdx = 0; respIdx < numRespondents; respIdx++) {
      const returnColIndex = respIdx + 2;
      row.getCell(respIdx + 1).value = { formula: `IFERROR(VLOOKUP("${fieldLabels[fieldIdx]}",Export!$B:$T,${returnColIndex},FALSE)," ")` };
    }
    row.commit();
  }

  for (let i = 1; i <= numRespondents; i++) sheet.getColumn(i).width = 18;
  sheet.views = [{ state: 'normal' }];
}

function createEstrazioneGraficiSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey, exportRange: string, numRespondents: number, fontName: string, headerArgb: string, sectionArgb: string): void {
  const sheet = workbook.addWorksheet('estrazione per grafici');
  const respondents = survey.respondents;
  const firstRespCol = 'D'; const lastRespCol = 'T';

  const headerRow = sheet.getRow(1);
  headerRow.getCell(1).value = 'Chiave'; headerRow.getCell(2).value = 'Domanda'; headerRow.getCell(3).value = 'MEDIE';
  respondents.forEach((r, idx) => { headerRow.getCell(4 + idx).value = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || ''; });
  styleHeaderRow(headerRow, fontName, headerArgb); headerRow.commit();

  let currentRow = 2;
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => { if (a === null) return 1; if (b === null) return -1; return a - b; });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];
    const blockRow = sheet.getRow(currentRow);
    blockRow.getCell(1).value = ''; blockRow.getCell(2).value = getBlockDisplayName(blockId);
    styleSectionRow(blockRow, fontName, sectionArgb); blockRow.commit(); currentRow++;

    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);
    sortedQuestions.forEach(question => {
      const row = sheet.getRow(currentRow);
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };
      const labelWithKey = question.questionKey ? `${question.questionKey} ${question.questionText}` : question.questionText;
      row.getCell(2).value = labelWithKey;
      row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

      if (question.type === 'scale_1_10_na') {
        row.getCell(3).value = { formula: `IFERROR(SUMIF(${firstRespCol}${currentRow}:${lastRespCol}${currentRow},">0")/COUNTIF(${firstRespCol}${currentRow}:${lastRespCol}${currentRow},">0")," ")` };
        row.getCell(3).font = { bold: true, name: fontName };
        row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
        row.getCell(3).alignment = { horizontal: 'center' };
      } else {
        row.getCell(3).value = '-';
        row.getCell(3).font = { italic: true, color: { argb: 'FF888888' }, name: fontName };
        row.getCell(3).alignment = { horizontal: 'center' };
      }

      for (let i = 0; i < numRespondents; i++) {
        row.getCell(4 + i).value = { formula: `IFERROR(VLOOKUP($A${currentRow},${exportRange},${3 + i},FALSE),"")` };
        row.getCell(4 + i).alignment = { horizontal: 'center' };
      }
      row.commit(); currentRow++;
    });
  });

  sheet.getColumn(1).width = 10; sheet.getColumn(2).width = 60; sheet.getColumn(3).width = 10;
  for (let i = 4; i <= 3 + numRespondents; i++) sheet.getColumn(i).width = 12;
  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
}

function createPerPdfSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey, exportRange: string, numRespondents: number, fontName: string, headerArgb: string, sectionArgb: string): void {
  const sheet = workbook.addWorksheet('per pdf');

  sheet.getRow(1).getCell(1).value = 'Riepilogo Survey';
  sheet.getRow(1).getCell(1).font = { bold: true, size: 16, name: fontName };
  sheet.getRow(1).height = 30;

  sheet.getRow(4).getCell(7).value = 'Seleziona partecipante:';
  sheet.getRow(4).getCell(7).font = { bold: true, name: fontName };
  sheet.getRow(5).getCell(8).value = 1;
  sheet.getRow(5).getCell(7).value = 'N°:';
  sheet.getCell('H5').font = { bold: true, size: 14, name: fontName };
  sheet.getCell('H5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
  sheet.getCell('H5').dataValidation = { type: 'whole', operator: 'between', formulae: [1, numRespondents], showErrorMessage: true, errorTitle: 'Valore non valido', error: `Inserire un numero da 1 a ${numRespondents}` };

  const metadataLabels = [
    { label: 'Cognome:', personeRow: 1 }, { label: 'Nome:', personeRow: 2 }, { label: 'Carica:', personeRow: 3 },
    { label: 'Comitato 1:', personeRow: 4 }, { label: 'Ruolo 1:', personeRow: 5 },
    { label: 'Comitato 2:', personeRow: 6 }, { label: 'Ruolo 2:', personeRow: 7 },
  ];

  metadataLabels.forEach((meta, idx) => {
    const row = sheet.getRow(8 + idx);
    row.getCell(6).value = meta.label; row.getCell(6).font = { bold: true, name: fontName }; row.getCell(6).alignment = { horizontal: 'right' };
    row.getCell(7).value = { formula: `INDEX(Persone!${meta.personeRow}:${meta.personeRow},'per pdf'!$H$5)` };
    row.getCell(7).alignment = { horizontal: 'left' };
  });

  let ifsArgs = '';
  for (let i = 1; i <= Math.min(numRespondents, 17); i++) {
    ifsArgs += `$H$5=${i},${2 + i}`;
    if (i < Math.min(numRespondents, 17)) ifsArgs += ',';
  }

  const hdrRow = sheet.getRow(16);
  hdrRow.getCell(1).value = 'Chiave'; hdrRow.getCell(2).value = 'Domanda'; hdrRow.getCell(3).value = 'MEDIE'; hdrRow.getCell(4).value = 'Risposta';
  styleHeaderRow(hdrRow, fontName, headerArgb);

  let currentRow = 17;
  const grouped = groupQuestionsByBlock(survey.questions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => { if (a === null) return 1; if (b === null) return -1; return a - b; });

  sortedBlocks.forEach(blockId => {
    const questions = grouped.get(blockId) || [];
    const blockRow = sheet.getRow(currentRow);
    blockRow.getCell(1).value = ''; blockRow.getCell(2).value = getBlockDisplayName(blockId);
    styleSectionRow(blockRow, fontName, sectionArgb); currentRow++;

    const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);
    sortedQuestions.forEach(question => {
      const row = sheet.getRow(currentRow);
      row.getCell(1).value = { formula: `IFERROR(LEFT(B${currentRow},SEARCH(" ",B${currentRow})-1),B${currentRow})` };
      const labelWithKey = question.questionKey ? `${question.questionKey} ${question.questionText}` : question.questionText;
      row.getCell(2).value = labelWithKey; row.getCell(2).alignment = { wrapText: true, vertical: 'top' };

      if (question.type === 'scale_1_10_na') {
        row.getCell(3).value = { formula: `IFERROR(VLOOKUP(A${currentRow},'estrazione per grafici'!$A:$C,3,FALSE),"")` };
        row.getCell(3).font = { bold: true, name: fontName };
        row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
        row.getCell(3).alignment = { horizontal: 'center' };
      } else {
        row.getCell(3).value = '-';
        row.getCell(3).font = { italic: true, color: { argb: 'FF888888' }, name: fontName };
        row.getCell(3).alignment = { horizontal: 'center' };
      }

      row.getCell(4).value = { formula: `IFERROR(VLOOKUP(A${currentRow},${exportRange},IFS(${ifsArgs}),FALSE),"")` };
      row.getCell(4).alignment = { horizontal: 'center' };
      currentRow++;
    });
  });

  sheet.getColumn(1).width = 10; sheet.getColumn(2).width = 60; sheet.getColumn(3).width = 10;
  sheet.getColumn(4).width = 20; sheet.getColumn(6).width = 12; sheet.getColumn(7).width = 25; sheet.getColumn(8).width = 8;
}

async function createGraficiSheet(workbook: ExcelJS.Workbook, survey: ParsedSurvey, template: ReturnType<typeof useTemplateStore.getState>['getActiveTemplate'] extends () => infer R ? R : never): Promise<void> {
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  if (scaleQuestions.length === 0) return;

  const fontName = template?.fontFamily || 'Calibri';
  const sheet = workbook.addWorksheet('Grafici');
  const Plotly = await import('plotly.js-dist-min');
  const container = document.createElement('div');
  container.style.position = 'absolute'; container.style.left = '-9999px';
  container.style.width = '900px'; container.style.height = '500px';
  document.body.appendChild(container);

  let currentRow = 1;

  try {
    const grouped = groupQuestionsByBlock(scaleQuestions);
    const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => { if (a === null) return 1; if (b === null) return -1; return a - b; });

    for (const blockId of sortedBlocks) {
      const questions = grouped.get(blockId) || [];
      if (questions.length === 0) continue;

      // Block label
      const labelRow = sheet.getRow(currentRow);
      labelRow.getCell(1).value = getBlockDisplayName(blockId);
      labelRow.getCell(1).font = { bold: true, size: 14, name: fontName };
      labelRow.commit();
      currentRow++;

      const chartHeight = Math.max(300, questions.length * 35 + 100);
      const meanPng = await generateBlockMeanChartPNG(blockId, questions, survey.scaleAnalytics, template, Plotly, container);
      const imageId = workbook.addImage({ base64: meanPng, extension: 'png' });
      sheet.addImage(imageId, { tl: { col: 0, row: currentRow }, ext: { width: 900, height: chartHeight } });
      currentRow += Math.ceil(chartHeight / 20) + 2;
    }
  } finally {
    Plotly.default.purge(container);
    document.body.removeChild(container);
  }
}

function styleHeaderRow(row: ExcelJS.Row, fontName: string, headerArgb: string): void {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: fontName };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerArgb } };
  row.alignment = { vertical: 'middle', horizontal: 'center' };
  row.height = 28;
}

function styleSectionRow(row: ExcelJS.Row, fontName: string, sectionArgb: string): void {
  row.font = { bold: true, italic: true, name: fontName };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sectionArgb } };
  row.height = 24;
}
