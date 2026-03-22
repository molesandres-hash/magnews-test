import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo } from '@/types/survey';
import { groupQuestionsByBlock, getSectionDisplayName } from './analytics';
import { useTemplateStore } from '@/store/templateStore';
import { hexToArgb } from './templateColors';
import { generateBlockMeanChartPNG, generateBlockDistributionChartPNG } from './excelChartHelper';
import { injectNativeCharts, type NativeChartDef } from './excelNativeCharts';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];
const COMITATO_PAGE_NAMES = ['Comitato 1', 'Comitato 2', 'Comitato 3'];
const DIST_COLORS = ['2563EB', 'DC2626', '16A34A', 'D97706', '7C3AED', 'DB2777', '0891B2', '65A30D'];

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

function cleanKey(raw: string | null): string {
  if (!raw) return '';
  return raw.replace(/^v/i, '').trim();
}

function extractGroupName(question: QuestionInfo): string | null {
  const m = question.cleanedHeader.match(/^\d+[\.\t\s]+(.+?)\s*[-–]\s*\d+\.\d+/);
  return m ? m[1].trim() : null;
}

export async function generateTabellaGrafici(
  survey: ParsedSurvey,
  mode: 'native' | 'png' = 'native'
): Promise<void> {
  const template = useTemplateStore.getState().getActiveTemplate();
  const fontName = template?.fontFamily || 'Calibri';
  const headerArgb = template ? hexToArgb(template.primaryColor) : 'FF2563EB';
  const blockHeaderArgb = template ? hexToArgb(template.secondaryColor) : 'FFE0E7FF';

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  let riepilogoSheet: ExcelJS.Worksheet | null = null;
  if (mode === 'native') {
    riepilogoSheet = workbook.addWorksheet('Riepilogo');
  }

  const sheet = workbook.addWorksheet('Foglio1');
  const respondents = survey.respondents;
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const numRespondents = respondents.length;

  // New column layout: A=key, B=text, C=MEDIE, D+=respondents, then counts
  const respondentStartCol = 4;
  const respondentEndCol = 3 + numRespondents;
  const countStartCol = respondentEndCol + 1;
  const totalCols = countStartCol + SCALE_ORDER.length - 1;

  const firstRespColLetter = colToLetter(respondentStartCol);
  const lastRespColLetter = colToLetter(respondentEndCol);

  // Logo
  if (template?.logoBase64) {
    const logoData = template.logoBase64.replace(/^data:image\/(png|jpeg|jpg);base64,/, '');
    const ext = template.logoBase64.includes('image/png') ? 'png' : 'jpeg';
    const imageId = workbook.addImage({ base64: logoData, extension: ext as 'png' | 'jpeg' });
    sheet.addImage(imageId, { tl: { col: totalCols - 2, row: 0 }, ext: { width: 120, height: 60 } });
  }

  // Header row
  const headerRow = sheet.getRow(1);
  headerRow.getCell(1).value = 'Chiave';
  headerRow.getCell(2).value = 'Domanda';
  headerRow.getCell(3).value = 'MEDIE';
  respondents.forEach((r, idx) => {
    const surname = r.originalData['Cognome'] || r.originalData['cognome'] || r.displayName.split(' ')[0] || r.displayName;
    headerRow.getCell(respondentStartCol + idx).value = surname;
  });
  SCALE_ORDER.forEach((scaleVal, idx) => {
    const cell = headerRow.getCell(countStartCol + idx);
    cell.value = scaleVal;
    cell.numFmt = '@';
  });
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: fontName };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerArgb } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  headerRow.height = 30;
  headerRow.commit();

  const grouped = groupQuestionsByBlock(scaleQuestions);
  const sortedBlocks = Array.from(grouped.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  // Plotly setup (PNG mode only)
  let Plotly: any = null;
  let container: HTMLDivElement | null = null;
  if (mode === 'png') {
    Plotly = await import('plotly.js-dist-min');
    container = document.createElement('div');
    container.style.position = 'absolute';
    container.style.left = '-9999px';
    container.style.width = '900px';
    container.style.height = '500px';
    document.body.appendChild(container);
  }

  const chartDefs: NativeChartDef[] = [];
  const sectionMeans: { name: string; mean: number }[] = [];
  const foglio1SheetIndex = mode === 'native' ? 2 : 1;

  let currentRow = 2;

  try {
    let lastSectionName: string | null = null;

    for (const blockId of sortedBlocks) {
      const questions = grouped.get(blockId) || [];
      const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);
      const sectionName = getSectionDisplayName(blockId, questions);
      const isComitato = COMITATO_PAGE_NAMES.includes(sectionName);

      // Spacer between comitato sections
      if (isComitato && lastSectionName && COMITATO_PAGE_NAMES.includes(lastSectionName)) {
        sheet.getRow(currentRow).height = 8;
        currentRow++;
      }

      // Section header (only when changes)
      if (sectionName !== lastSectionName) {
        lastSectionName = sectionName;
        const blockRow = sheet.getRow(currentRow);
        blockRow.getCell(1).value = sectionName;
        blockRow.font = { bold: true, size: 12, name: fontName };
        blockRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: blockHeaderArgb } };
        blockRow.height = 28;
        sheet.mergeCells(currentRow, 1, currentRow, totalCols);
        blockRow.commit();
        currentRow++;
      }

      // Group sub-headers tracking
      let lastGroupName: string | null = null;
      const dataStartRow = currentRow;

      for (const question of sortedQuestions) {
        const analytics = survey.scaleAnalytics.get(question.id);
        if (!analytics) continue;

        // Group sub-header
        const groupName = extractGroupName(question);
        if (groupName && groupName !== lastGroupName) {
          lastGroupName = groupName;
          const grpRow = sheet.getRow(currentRow);
          sheet.mergeCells(currentRow, 1, currentRow, totalCols);
          grpRow.getCell(1).value = `  ${groupName}`;
          grpRow.getCell(1).font = { bold: true, size: 10, color: { argb: 'FF1E3A5F' }, name: fontName, italic: true };
          grpRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
          grpRow.getCell(1).alignment = { horizontal: 'left' };
          grpRow.height = 17;
          grpRow.commit();
          currentRow++;
        }

        const dataRow = sheet.getRow(currentRow);

        // Col A: question key
        dataRow.getCell(1).value = cleanKey(question.questionKey);
        dataRow.getCell(1).font = { bold: true, size: 9, color: { argb: 'FF2563EB' }, name: fontName };
        dataRow.getCell(1).alignment = { horizontal: 'center' };

        // Col B: clean question text
        const displayText = question.questionText.length > 90
          ? question.questionText.slice(0, 87) + '...'
          : question.questionText;
        dataRow.getCell(2).value = displayText;
        dataRow.getCell(2).alignment = { wrapText: false, vertical: 'middle' };
        dataRow.getCell(2).font = { size: 9, name: fontName };
        if (question.questionText.length > 90) {
          dataRow.getCell(2).note = { texts: [{ font: { size: 9 }, text: question.questionText }] };
        }

        // Col C: MEDIE formula
        dataRow.getCell(3).value = {
          formula: `IFERROR(SUMIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},">0")/COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},">0")," ")`
        };
        dataRow.getCell(3).alignment = { horizontal: 'center' };
        dataRow.getCell(3).font = { bold: true, name: fontName };
        dataRow.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };

        // Respondent values
        respondents.forEach((r, idx) => {
          const value = analytics.respondentValues[r.id];
          const cell = dataRow.getCell(respondentStartCol + idx);
          if (value === null || value === undefined) {
            cell.value = value === undefined ? '—' : 'n.r.';
            cell.font = value === undefined
              ? { color: { argb: 'FFCCCCCC' }, name: fontName }
              : { italic: true, color: { argb: 'FF888888' }, name: fontName };
          } else {
            cell.value = value;
            if (value >= 8) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            else if (value >= 5) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
            else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          }
          cell.alignment = { horizontal: 'center' };
        });

        // Count columns
        SCALE_ORDER.forEach((scaleVal, idx) => {
          const cell = dataRow.getCell(countStartCol + idx);
          if (scaleVal === 'N/A') {
            cell.value = { formula: `COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},"n.r.")` };
          } else {
            cell.value = { formula: `COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},${scaleVal})` };
          }
          cell.alignment = { horizontal: 'center' };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
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
      }

      const dataEndRow = currentRow - 1;
      const questionsWithAnalytics = sortedQuestions.filter(q => survey.scaleAnalytics.get(q.id));

      // Charts per section
      if (questionsWithAnalytics.length > 0) {
        if (mode === 'native') {
          const means = questionsWithAnalytics.map(q => survey.scaleAnalytics.get(q.id)!.mean);
          const grandMean = means.reduce((a, b) => a + b, 0) / means.length;
          sectionMeans.push({ name: sectionName, mean: grandMean });

          currentRow += 2;
          const meanChartRows = Math.max(15, questionsWithAnalytics.length * 2 + 5);

          // Mean chart — catRef=col B (short text), valRef=col C (MEDIE)
          chartDefs.push({
            sheetIndex: foglio1SheetIndex,
            sheetName: sheet.name,
            title: sectionName,
            direction: 'horizontal',
            anchor: { fromRow: currentRow - 1, fromCol: 0, toRow: currentRow + meanChartRows - 1, toCol: 13 },
            series: [{
              name: 'Media',
              catRef: `'${sheet.name}'!$B$${dataStartRow}:$B$${dataEndRow}`,
              valRef: `'${sheet.name}'!$C$${dataStartRow}:$C$${dataEndRow}`,
              color: template?.primaryColor?.replace('#', '') ?? '2563EB',
            }],
            valAxisMin: 0,
            valAxisMax: 10,
          });
          currentRow += meanChartRows + 2;

          // Distribution chart
          const distChartRows = 15;
          const cStartLetter = colToLetter(countStartCol);
          const cEndLetter = colToLetter(countStartCol + 10);

          chartDefs.push({
            sheetIndex: foglio1SheetIndex,
            sheetName: sheet.name,
            title: `${sectionName} — Distribuzione`,
            direction: 'vertical',
            anchor: { fromRow: currentRow - 1, fromCol: 0, toRow: currentRow + distChartRows - 1, toCol: 13 },
            series: questionsWithAnalytics.map((q, qIdx) => ({
              name: cleanKey(q.questionKey) || `Q${qIdx + 1}`,
              catRef: `'${sheet.name}'!$${cStartLetter}$1:$${cEndLetter}$1`,
              valRef: `'${sheet.name}'!$${cStartLetter}$${dataStartRow + qIdx}:$${cEndLetter}$${dataStartRow + qIdx}`,
              color: DIST_COLORS[qIdx % DIST_COLORS.length],
            })),
          });
          currentRow += distChartRows + 3;

        } else {
          // PNG mode (existing behavior)
          const meanPng = await generateBlockMeanChartPNG(blockId, questions, survey.scaleAnalytics, template, Plotly, container!);
          const meanChartHeight = Math.max(300, questions.length * 35 + 100);
          const meanImageId = workbook.addImage({ base64: meanPng, extension: 'png' });
          sheet.addImage(meanImageId, { tl: { col: 0, row: currentRow }, ext: { width: 900, height: meanChartHeight } });
          currentRow += Math.ceil(meanChartHeight / 20) + 2;

          const distPng = await generateBlockDistributionChartPNG(blockId, questions, survey.scaleAnalytics, template, Plotly, container!);
          const distImageId = workbook.addImage({ base64: distPng, extension: 'png' });
          sheet.addImage(distImageId, { tl: { col: 0, row: currentRow }, ext: { width: 900, height: 400 } });
          currentRow += Math.ceil(400 / 20) + 2;
        }
      }
    }
  } finally {
    if (mode === 'png' && Plotly && container) {
      Plotly.default.purge(container);
      document.body.removeChild(container);
    }
  }

  // Column widths — new layout
  sheet.getColumn(1).width = 6;
  sheet.getColumn(2).width = 52;
  sheet.getColumn(3).width = 10;
  for (let i = respondentStartCol; i <= respondentEndCol; i++) sheet.getColumn(i).width = 14;
  for (let i = countStartCol; i <= totalCols; i++) sheet.getColumn(i).width = 6;
  sheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 1 }];

  // Riepilogo sheet content (native mode)
  if (mode === 'native' && riepilogoSheet && sectionMeans.length > 0) {
    sectionMeans.sort((a, b) => b.mean - a.mean);

    riepilogoSheet.getRow(1).getCell(1).value = `Media per sezione — ${survey.metadata.fileName.replace('.csv', '')}`;
    riepilogoSheet.getRow(1).font = { bold: true, size: 14, name: fontName };
    riepilogoSheet.getColumn(1).width = 40;
    riepilogoSheet.getColumn(2).width = 10;

    const dataStart = 100;
    const noteRow = riepilogoSheet.getRow(dataStart - 1);
    noteRow.getCell(1).value = 'Dati riepilogo (non modificare)';
    noteRow.getCell(1).font = { color: { argb: 'FF999999' }, size: 8, name: fontName };

    sectionMeans.forEach((sm, i) => {
      const row = riepilogoSheet!.getRow(dataStart + i);
      row.getCell(1).value = sm.name;
      row.getCell(2).value = Math.round(sm.mean * 100) / 100;
      row.font = { color: { argb: 'FF999999' }, size: 9, name: fontName };
    });

    const dataEnd = dataStart + sectionMeans.length - 1;

    chartDefs.push({
      sheetIndex: 1,
      sheetName: 'Riepilogo',
      title: 'Media per sezione',
      direction: 'horizontal',
      anchor: { fromRow: 2, fromCol: 0, toRow: 22, toCol: 10 },
      series: [{
        name: 'Media',
        catRef: `'Riepilogo'!$A$${dataStart}:$A$${dataEnd}`,
        valRef: `'Riepilogo'!$B$${dataStart}:$B$${dataEnd}`,
        color: template?.primaryColor?.replace('#', '') ?? '2563EB',
      }],
      valAxisMin: 0,
      valAxisMax: 10,
    });
  }

  // Write buffer
  const buffer = await workbook.xlsx.writeBuffer();

  let finalBuffer: ArrayBuffer;
  if (mode === 'native' && chartDefs.length > 0) {
    try {
      finalBuffer = await injectNativeCharts(buffer as ArrayBuffer, chartDefs);
    } catch (e) {
      console.warn('Native chart injection failed, saving without charts:', e);
      finalBuffer = buffer as ArrayBuffer;
    }
  } else {
    finalBuffer = buffer as ArrayBuffer;
  }

  const blob = new Blob([finalBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fileName = mode === 'native'
    ? `tabella_grafici_NATIVA_${survey.metadata.fileName.replace('.csv', '')}.xlsx`
    : `tabella_grafici_PNG_${survey.metadata.fileName.replace('.csv', '')}.xlsx`;
  saveAs(blob, fileName);
}
