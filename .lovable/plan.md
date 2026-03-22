

## Plan: Fix Native Charts + Restructure Tabella Grafici Layout

### Problem Analysis

**Native charts issues (from code review):**
1. Chart category labels reference column A (`$A$dataStartRow:$A$dataEndRow`) which contains full question text — extremely long strings that render illegibly in Excel chart axes
2. The mean chart `valRef` references column B (MEDIE), but MEDIE is an IFERROR formula. When Excel opens the file fresh, these formula cells may not yet have cached values, causing the chart to show nothing or errors
3. `respondentValues[r.id]` can be `undefined` (not just `null`) — the code only checks `=== null`, so `undefined` values get written as `cell.value = undefined`, potentially corrupting cell data
4. Distribution chart `catRef` references row 1 headers which are set with `numFmt: '@'` (text format) — this can cause issues with Excel's auto-detection of category vs numeric data

**Layout issues:**
- Column A shows full question text, no key/sub-key separation
- No group sub-headers within sections
- Question text not cleaned (shows full raw header with group prefix)

---

### Changes (all in `tabellaGraficiWriter.ts` only)

#### 1. Add helper functions

```ts
function extractGroupName(question: QuestionInfo): string | null {
  const m = question.cleanedHeader.match(/^\d+[\.\t\s]+(.+?)\s*[-–]\s*\d+\.\d+/);
  return m ? m[1].trim() : null;
}

function cleanKey(raw: string | null): string {
  if (!raw) return '';
  return raw.replace(/^v/i, '').trim();
}
```

#### 2. Restructure column layout

- **Column A** (width 6): question key only (`cleanKey(question.questionKey)`)
- **Column B** (width 52): clean question text (`question.questionText`, truncated to 90 chars)
- **Column C**: MEDIE formula (same as current column B)
- Respondent columns shift to start at column 4 (instead of 3)
- Count columns shift accordingly
- Header row updated: cell 1 = "Chiave", cell 2 = "Domanda", cell 3 = "MEDIE"

#### 3. Add group sub-headers

Inside the block loop, track `lastGroupName`. When group changes (extracted via regex from `cleanedHeader`), write a styled sub-header row:
- Merged across all columns
- Italic bold, blue text on light indigo background
- Height 17px

#### 4. Fix native chart references

After restructuring columns, update chart series references:
- Mean chart `catRef`: reference column A (short keys like "5.1") or column B (question text) — use column B for readability but with shorter text now
- Mean chart `valRef`: reference column C (MEDIE)
- Distribution chart: update `countStartCol` references to match new column offsets
- Add data labels to chart XML (`<c:dLbls>` element) so values are visible on bars

#### 5. Fix undefined respondent values

Change the respondent value check from:
```ts
if (value === null) {
```
to:
```ts
if (value === null || value === undefined) {
```

This prevents `undefined` from being written as cell values.

#### 6. Improve chart XML quality

In `excelNativeCharts.ts`, enhance `buildChartXml`:
- Add `<c:dLbls>` (data labels) to show values on bars
- Add `<c:overlap val="-10"/>` for better bar spacing
- For horizontal charts with many categories, increase chart height calculation
- Use unique axis IDs per chart (currently all use 111/222 which is fine since each is in a separate XML file, but good practice)

#### 7. Riepilogo chart fix

The Riepilogo chart references hidden data rows (100+). This should work correctly — verify the `catRef` and `valRef` point to the right cells after column restructure. Since Riepilogo has its own independent layout, it's unaffected by the Foglio1 column changes.

---

### Summary of column mapping changes

| Before | After |
|--------|-------|
| Col A = question text (full) | Col A = question key ("5.1") |
| Col B = MEDIE | Col B = question text (clean, 90 chars) |
| Col C+ = respondents | Col C = MEDIE |
| After respondents = counts | Col D+ = respondents, then counts |

### Files modified
- `src/utils/tabellaGraficiWriter.ts` — layout restructure + chart ref fixes
- `src/utils/excelNativeCharts.ts` — add data labels to chart XML

