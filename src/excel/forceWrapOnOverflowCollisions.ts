import type { Cell, Sheet } from '@fortune-sheet/core';

const DEFAULT_COLLEN = 72;
// Cell horizontal padding budget that fortune-sheet reserves for borders +
// breathing room; subtract from column width when comparing to text width.
const CELL_HPADDING = 4;

let sharedCtx: CanvasRenderingContext2D | null = null;
const getMeasureCtx = (): CanvasRenderingContext2D | null => {
  if (sharedCtx) return sharedCtx;
  const canvas = typeof document !== 'undefined' ? document.createElement('canvas') : null;
  sharedCtx = canvas ? canvas.getContext('2d') : null;
  return sharedCtx;
};

const extractText = (v: Cell): string => {
  if (typeof v.v === 'string') return v.v;
  if (typeof v.v === 'number') return String(v.v);
  if (typeof v.m === 'string') return v.m;
  if (typeof v.m === 'number') return String(v.m);
  const ct = v.ct as { t?: string; s?: Array<{ v?: unknown }> } | undefined;
  if (ct?.t === 'inlineStr' && Array.isArray(ct.s)) {
    return ct.s
      .map((run) => (run?.v == null ? '' : String(run.v)))
      .join('');
  }
  return '';
};

const hasContent = (v: Cell | null | undefined): boolean => {
  if (!v) return false;
  return extractText(v).length > 0;
};

const measureText = (text: string, fontSize: number, fontFamily: string): number => {
  const ctx = getMeasureCtx();
  if (!ctx) {
    // Fallback heuristic if no canvas is available (SSR, unlikely here).
    return text.length * fontSize * 0.6;
  }
  ctx.font = `${fontSize}px "${fontFamily}", Arial, sans-serif`;
  return ctx.measureText(text).width;
};

export const forceWrapOnOverflowCollisions = (sheets: Sheet[]): Sheet[] => {
  return sheets.map((sheet) => {
    if (!sheet.celldata || sheet.celldata.length === 0) return sheet;

    const nextCelldata = sheet.celldata.map((entry) => ({
      ...entry,
      v: entry.v ? { ...entry.v } : entry.v,
    }));

    const cellByKey = new Map<string, Cell | null | undefined>();
    for (const entry of nextCelldata) {
      cellByKey.set(`${entry.r}:${entry.c}`, entry.v ?? null);
    }

    const columnlen = sheet.config?.columnlen ?? {};

    let mutated = false;
    for (const entry of nextCelldata) {
      const v = entry.v;
      if (!v) continue;
      if (v.tb !== '1') continue;
      if (v.mc) continue;

      const text = extractText(v);
      if (text.length === 0) continue;

      const fontSize = typeof v.fs === 'number' ? v.fs : 11;
      const fontFamily = typeof v.ff === 'string' ? v.ff : 'Arial';
      const colWidth =
        typeof columnlen[String(entry.c)] === 'number'
          ? columnlen[String(entry.c)]
          : DEFAULT_COLLEN;
      const textWidth = measureText(text, fontSize, fontFamily);
      if (textWidth <= colWidth - CELL_HPADDING) continue;

      const ht = typeof v.ht === 'number' ? v.ht : 1;
      const neighbors =
        ht === 0 ? [entry.c - 1, entry.c + 1] : ht === 2 ? [entry.c - 1] : [entry.c + 1];

      const collides = neighbors.some((n) =>
        n >= 0 ? hasContent(cellByKey.get(`${entry.r}:${n}`)) : false,
      );
      if (!collides) continue;

      v.tb = '2';
      mutated = true;
    }

    if (!mutated) return sheet;
    return { ...sheet, celldata: nextCelldata };
  });
};
