import type { Sheet } from '@fortune-sheet/core';

type CellV = NonNullable<NonNullable<Sheet['celldata']>[number]['v']>;

// Concatenate the text of an inlineStr cell's rich-text runs.
const extractInlineStrText = (v: CellV): string | null => {
  const ct = v.ct as { t?: string; s?: Array<{ v?: unknown }> } | undefined;
  if (ct?.t !== 'inlineStr' || !Array.isArray(ct.s) || ct.s.length === 0) {
    return null;
  }
  let text = '';
  for (const run of ct.s) {
    if (run?.v != null) text += String(run.v);
  }
  return text.length > 0 ? text : null;
};

// fortune-sheet's `cellOverflow_trace` (core:73721) decides whether a
// neighbor blocks overflow by checking `!_.isEmpty(cell.v)`. For cells that
// fortune-excel emits as inlineStr rich text, `cell.v` is left undefined —
// the text lives in `cell.ct.s[i].v`. fortune-sheet's overflow trace then
// treats such cells as empty and happily pushes the source cell's text
// across them, producing the "merged with neighbor" rendering we see on
// headers like Q3 whose neighbor P3 is inlineStr.
//
// The cell-rendering path for the cell itself already guards with
// `isInlineStringCell(cell)` (core:73200), so the source never mis-renders.
// The fix is to backfill `v.v` with the concatenated run text so the trace
// check matches what `isInlineStringCell` would say. Rich-text rendering
// continues to come from `ct.s`, so no visual side effect on the cell
// itself.
export const backfillInlineStringV = (sheets: Sheet[]): Sheet[] => {
  return sheets.map((sheet) => {
    if (!sheet.celldata || sheet.celldata.length === 0) return sheet;

    let mutated = false;
    const nextCelldata = sheet.celldata.map((entry) => {
      const v = entry.v;
      if (!v) return entry;
      // If v.v is already a non-empty string/number, nothing to do.
      if (typeof v.v === 'string' && v.v.length > 0) return entry;
      if (typeof v.v === 'number') return entry;
      const text = extractInlineStrText(v);
      if (text == null) return entry;
      mutated = true;
      return { ...entry, v: { ...v, v: text } };
    });

    if (!mutated) return sheet;
    return { ...sheet, celldata: nextCelldata };
  });
};
