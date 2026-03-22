import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { ParsedSurvey } from '@/types/survey';
import { groupQuestionsByBlock, getSectionDisplayName } from './analytics';
import { useTemplateStore } from '@/store/templateStore';
import { hexToArgb } from './templateColors';
import { generateBlockMeanChartPNG, generateBlockDistributionChartPNG } from './excelChartHelper';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];
const COMITATO_PAGE_NAMES = ['Comitato 1', 'Comitato 2', 'Comitato 3'];

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

export async function generateTabellaGrafici(survey: ParsedSurvey): Promise<void> {
  const template = useTemplateStore.getState().getActiveTemplate();
  const fontName = template?.fontFamily || 'Calibri';
  const headerArgb = template ? hexToArgb(template.primaryColor) : 'FF2563EB';
  const blockHeaderArgb = template ? hexToArgb(template.secondaryColor) : 'FFE0E7FF';

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Magnews Survey Analyzer';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Foglio1');
  const respondents = survey.respondents;
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const numRespondents = respondents.length;

  const respondentStartCol = 3;
  const respondentEndCol = 2 + numRespondents;
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
  headerRow.getCell(1).value = 'Domanda';
  headerRow.getCell(2).value = 'MEDIE';
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

  // Plotly setup for chart embedding
  const Plotly = await import('plotly.js-dist-min');
  const container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.width = '900px';
  container.style.height = '500px';
  document.body.appendChild(container);

  let currentRow = 2;

  try {
    let lastSectionName: string | null = null;

    for (const blockId of sortedBlocks) {
      const questions = grouped.get(blockId) || [];
      const sortedQuestions = [...questions].sort((a, b) => a.subId - b.subId);

      // Determine if this block contains comitato questions that need visual separation
      const sectionName = getSectionDisplayName(blockId, questions);
      const isComitato = COMITATO_PAGE_NAMES.includes(sectionName);

      // Add spacer + divider between comitato sections
      if (isComitato && lastSectionName && COMITATO_PAGE_NAMES.includes(lastSectionName)) {
        sheet.getRow(currentRow).height = 8;
        currentRow++;
      }

      // Write section header only when section changes
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

      for (const question of sortedQuestions) {
        const analytics = survey.scaleAnalytics.get(question.id);
        if (!analytics) continue;

        const dataRow = sheet.getRow(currentRow);
        dataRow.getCell(1).value = question.questionText;
        dataRow.getCell(1).alignment = { wrapText: true, vertical: 'top' };

        dataRow.getCell(2).value = {
          formula: `IFERROR(SUMIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},">0")/COUNTIF(${firstRespColLetter}${currentRow}:${lastRespColLetter}${currentRow},">0")," ")`
        };
        dataRow.getCell(2).alignment = { horizontal: 'center' };
        dataRow.getCell(2).font = { bold: true, name: fontName };
        dataRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };

        respondents.forEach((r, idx) => {
          const value = analytics.respondentValues[r.id];
          const cell = dataRow.getCell(respondentStartCol + idx);
          if (value === null) {
            cell.value = 'n.r.';
            cell.font = { italic: true, color: { argb: 'FF888888' }, name: fontName };
          } else {
            cell.value = value;
            if (value >= 8) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            else if (value >= 5) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
            else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          }
          cell.alignment = { horizontal: 'center' };
        });

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

      // Embed mean chart for this block
      if (questions.length > 0) {
        const meanPng = await generateBlockMeanChartPNG(blockId, questions, survey.scaleAnalytics, template, Plotly, container);
        const meanChartHeight = Math.max(300, questions.length * 35 + 100);
        const meanImageId = workbook.addImage({ base64: meanPng, extension: 'png' });
        sheet.addImage(meanImageId, { tl: { col: 0, row: currentRow }, ext: { width: 900, height: meanChartHeight } });
        currentRow += Math.ceil(meanChartHeight / 20) + 2;

        // Embed distribution chart
        const distPng = await generateBlockDistributionChartPNG(blockId, questions, survey.scaleAnalytics, template, Plotly, container);
        const distImageId = workbook.addImage({ base64: distPng, extension: 'png' });
        sheet.addImage(distImageId, { tl: { col: 0, row: currentRow }, ext: { width: 900, height: 400 } });
        currentRow += Math.ceil(400 / 20) + 2;
      }
    }
  } finally {
    Plotly.default.purge(container);
    document.body.removeChild(container);
  }

  // Column widths
  sheet.getColumn(1).width = 60;
  sheet.getColumn(2).width = 10;
  for (let i = respondentStartCol; i <= respondentEndCol; i++) sheet.getColumn(i).width = 14;
  for (let i = countStartCol; i <= totalCols; i++) sheet.getColumn(i).width = 6;
  sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, 'tabella_grafici_scala10_NA_GENERATA.xlsx');
}
