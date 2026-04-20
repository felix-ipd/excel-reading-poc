import type { Sheet } from '@fortune-sheet/core';

// fortune-sheet defaults from @fortune-sheet/core/dist/index.esm.js:63380-63387
const ROW_HEADER_WIDTH = 46;
const COLUMN_HEADER_HEIGHT = 20;
const DEFAULT_COLLEN = 73;
const DEFAULT_ROWLEN = 19;

// Height of fortune-stat-area at the bottom (the selection calc-info bar).
// See .luckysheet-sheet-selection-calInfo in index.css.
const STAT_AREA_HEIGHT = 22;

// Tiny slack for borders + potential scrollbar gutter.
const CHROME_SLACK = 4;

// Mirrors the per-row / per-col computation in core's `calcRowColSize`
// (index.esm.js:63469): each row/column contributes `round((len + 1) * zoom)`.
// Zoom ratio can be baked into the sheet (xlsx `<sheetView zoomScale="...">`).
export const measureSheet = (sheet: Sheet): { width: number; height: number } => {
  const columnlen = sheet.config?.columnlen ?? {};
  const rowlen = sheet.config?.rowlen ?? {};
  const columns = sheet.column ?? 0;
  const rows = sheet.row ?? 0;
  const zoom = sheet.zoomRatio ?? 1;

  let width = Math.round(ROW_HEADER_WIDTH * zoom);
  for (let c = 0; c < columns; c++) {
    const raw = columnlen[String(c)];
    const w = typeof raw === 'number' ? raw : DEFAULT_COLLEN;
    width += Math.round((w + 1) * zoom);
  }

  let height = Math.round(COLUMN_HEADER_HEIGHT * zoom);
  for (let r = 0; r < rows; r++) {
    const raw = rowlen[String(r)];
    const h = typeof raw === 'number' ? raw : DEFAULT_ROWLEN;
    height += Math.round((h + 1) * zoom);
  }
  height += STAT_AREA_HEIGHT;

  return {
    width: width + CHROME_SLACK,
    height: height + CHROME_SLACK,
  };
};
