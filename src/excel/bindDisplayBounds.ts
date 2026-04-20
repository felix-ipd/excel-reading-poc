import type { Sheet } from '@fortune-sheet/core';

type MergeMap = NonNullable<NonNullable<Sheet['config']>['merge']>;

const foldMergesIntoMax = (
  merges: MergeMap | undefined,
  runningMaxR: number,
  runningMaxC: number,
): { maxR: number; maxC: number } => {
  if (!merges) return { maxR: runningMaxR, maxC: runningMaxC };
  let maxR = runningMaxR;
  let maxC = runningMaxC;
  for (const entry of Object.values(merges)) {
    if (!entry) continue;
    const r2 = entry.r + (entry.rs ?? 1) - 1;
    const c2 = entry.c + (entry.cs ?? 1) - 1;
    if (r2 > maxR) maxR = r2;
    if (c2 > maxC) maxC = c2;
  }
  return { maxR, maxC };
};

export const bindDisplayBounds = (sheets: Sheet[]): Sheet[] => {
  return sheets.map((sheet) => {
    let maxR = -1;
    let maxC = -1;
    if (sheet.celldata) {
      for (const entry of sheet.celldata) {
        if (entry.r > maxR) maxR = entry.r;
        if (entry.c > maxC) maxC = entry.c;
      }
    }
    const folded = foldMergesIntoMax(sheet.config?.merge, maxR, maxC);
    maxR = folded.maxR;
    maxC = folded.maxC;

    // +2 rather than +1 so there's one empty buffer row/column past the
    // data — gives the reader a visual hint that the content ends here.
    const row = maxR >= 0 ? maxR + 2 : 1;
    const column = maxC >= 0 ? maxC + 2 : 1;

    return { ...sheet, row, column, addRows: 0 };
  });
};
