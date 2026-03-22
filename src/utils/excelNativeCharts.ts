import JSZip from 'jszip';

export interface NativeChartDef {
  sheetIndex: number; // 1-based (sheet1.xml = 1)
  sheetName: string;
  title: string;
  direction: 'horizontal' | 'vertical';
  anchor: { fromRow: number; fromCol: number; toRow: number; toCol: number };
  series: Array<{
    name: string;
    catRef: string;  // e.g., "'Foglio1'!$A$3:$A$10"
    valRef: string;
    color: string;   // hex without #
  }>;
  valAxisMin?: number;
  valAxisMax?: number;
}

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function buildChartXml(chart: NativeChartDef): string {
  const dir = chart.direction === 'horizontal' ? 'bar' : 'col';
  const catPos = chart.direction === 'horizontal' ? 'l' : 'b';
  const valPos = chart.direction === 'horizontal' ? 'b' : 'l';

  const serXml = chart.series.map((s, i) =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
    `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>` +
    `<c:spPr><a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill></c:spPr>` +
    `<c:cat><c:strRef><c:f>${esc(s.catRef)}</c:f></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>${esc(s.valRef)}</c:f></c:numRef></c:val>` +
    `</c:ser>`
  ).join('');

  const scaling = [
    '<c:scaling><c:orientation val="minMax"/>',
    chart.valAxisMax != null ? `<c:max val="${chart.valAxisMax}"/>` : '',
    chart.valAxisMin != null ? `<c:min val="${chart.valAxisMin}"/>` : '',
    '</c:scaling>'
  ].join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
<c:chart>
<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="it-IT" sz="1200" b="1"/><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/>
<c:plotArea><c:layout/>
<c:barChart><c:barDir val="${dir}"/><c:grouping val="clustered"/><c:varyColors val="0"/>
${serXml}
<c:axId val="111"/><c:axId val="222"/></c:barChart>
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:crossAx val="222"/></c:catAx>
<c:valAx><c:axId val="222"/>${scaling}<c:delete val="0"/><c:axPos val="${valPos}"/><c:numFmt formatCode="0.0" sourceLinked="0"/><c:crossAx val="111"/><c:majorGridlines/></c:valAx>
</c:plotArea>
<c:legend><c:legendPos val="b"/></c:legend>
<c:plotVisOnly val="1"/>
</c:chart>
</c:chartSpace>`;
}

function buildAnchorXml(rId: string, anchor: NativeChartDef['anchor'], cNvPrId: number): string {
  return `<xdr:twoCellAnchor>
<xdr:from><xdr:col>${anchor.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>${anchor.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro="">
<xdr:nvGraphicFramePr><xdr:cNvPr id="${cNvPrId}" name="Chart ${cNvPrId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rId}"/></a:graphicData></a:graphic>
</xdr:graphicFrame>
<xdr:clientData/>
</xdr:twoCellAnchor>`;
}

function buildDrawingXml(anchors: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
${anchors.join('\n')}
</xdr:wsDr>`;
}

export async function injectNativeCharts(
  xlsxBuffer: ArrayBuffer,
  charts: NativeChartDef[]
): Promise<ArrayBuffer> {
  if (charts.length === 0) return xlsxBuffer;

  const zip = await JSZip.loadAsync(xlsxBuffer);
  let contentTypes = await zip.file('[Content_Types].xml')!.async('string');

  // Group charts by sheet
  const bySheet = new Map<number, NativeChartDef[]>();
  for (const chart of charts) {
    const list = bySheet.get(chart.sheetIndex) || [];
    list.push(chart);
    bySheet.set(chart.sheetIndex, list);
  }

  let globalChartNum = 1;

  for (const [sheetIdx, sheetCharts] of bySheet) {
    const sheetFile = `xl/worksheets/sheet${sheetIdx}.xml`;
    const sheetRelsFile = `xl/worksheets/_rels/sheet${sheetIdx}.xml.rels`;

    let sheetXml = await zip.file(sheetFile)?.async('string');
    if (!sheetXml) continue;

    // Check for existing drawing in sheet
    const existingDrawingMatch = sheetXml.match(/<drawing r:id="([^"]+)"/);
    let existingDrawingFile: string | null = null;
    let existingDrawingRelsFile: string | null = null;

    if (existingDrawingMatch) {
      const drawingRId = existingDrawingMatch[1];
      const sheetRelsXml = await zip.file(sheetRelsFile)?.async('string');
      if (sheetRelsXml) {
        const match = sheetRelsXml.match(new RegExp(`Id="${drawingRId}"[^>]*Target="([^"]+)"`));
        if (match) {
          const target = match[1].replace(/^\.\.\//g, '');
          existingDrawingFile = target.startsWith('xl/') ? target : `xl/${target}`;
          if (existingDrawingFile.startsWith('xl/xl/')) existingDrawingFile = existingDrawingFile.replace('xl/xl/', 'xl/');
          existingDrawingRelsFile = existingDrawingFile.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        }
      }
    }

    // Build chart files and anchor XML
    const anchorXmls: string[] = [];
    const chartRels: { rId: string; target: string }[] = [];

    for (let i = 0; i < sheetCharts.length; i++) {
      const chart = sheetCharts[i];
      const chartNum = globalChartNum++;
      const chartFile = `xl/charts/chart${chartNum}.xml`;
      const rId = `rIdNChart${chartNum}`;

      zip.file(chartFile, buildChartXml(chart));
      contentTypes = contentTypes.replace(
        '</Types>',
        `<Override PartName="/${chartFile}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`
      );

      anchorXmls.push(buildAnchorXml(rId, chart.anchor, 1000 + chartNum));
      chartRels.push({ rId, target: `../charts/chart${chartNum}.xml` });
    }

    if (existingDrawingFile) {
      // Append to existing drawing
      let drawingXml = await zip.file(existingDrawingFile)!.async('string');
      drawingXml = drawingXml.replace('</xdr:wsDr>', anchorXmls.join('\n') + '\n</xdr:wsDr>');
      zip.file(existingDrawingFile, drawingXml);

      // Update drawing rels
      let drawingRelsXml = await zip.file(existingDrawingRelsFile!)?.async('string') ||
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_REL}"></Relationships>`;
      for (const rel of chartRels) {
        drawingRelsXml = drawingRelsXml.replace(
          '</Relationships>',
          `<Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${rel.target}"/></Relationships>`
        );
      }
      zip.file(existingDrawingRelsFile!, drawingRelsXml);
    } else {
      // Create new drawing
      const drawingNum = sheetIdx + 100;
      const drawingFile = `xl/drawings/drawing${drawingNum}.xml`;
      const drawingRelsFile = `xl/drawings/_rels/drawing${drawingNum}.xml.rels`;

      zip.file(drawingFile, buildDrawingXml(anchorXmls));

      // Drawing rels
      let drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_REL}">`;
      for (const rel of chartRels) {
        drawingRelsXml += `<Relationship Id="${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${rel.target}"/>`;
      }
      drawingRelsXml += '</Relationships>';
      zip.file(drawingRelsFile, drawingRelsXml);

      // Content type for drawing
      contentTypes = contentTypes.replace(
        '</Types>',
        `<Override PartName="/${drawingFile}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
      );

      // Add drawing ref to sheet
      const drawingRId = `rIdDrw${drawingNum}`;
      let sheetRelsXml = await zip.file(sheetRelsFile)?.async('string') ||
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_REL}"></Relationships>`;
      sheetRelsXml = sheetRelsXml.replace(
        '</Relationships>',
        `<Relationship Id="${drawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNum}.xml"/></Relationships>`
      );
      zip.file(sheetRelsFile, sheetRelsXml);

      // Add <drawing> to sheet XML — insert before </worksheet>
      sheetXml = sheetXml.replace('</worksheet>', `<drawing r:id="${drawingRId}"/></worksheet>`);
      zip.file(sheetFile, sheetXml);
    }
  }

  zip.file('[Content_Types].xml', contentTypes);
  return await zip.generateAsync({ type: 'arraybuffer' });
}
