import type { Cell, Sheet } from '@fortune-sheet/core';
import { loadXlsx } from './xlsxZip';
import { parseTheme } from './theme';
import { parseDxfs, type Dxf, type BorderSide } from './dxf';
import { parseTableStyles, type TableElementType } from './tableStyles';
import { indexSheetTables, parseTable, type CellRange, type TableDef } from './tables';
import { resolveColor, type ThemePalette } from './colors';

type EffectiveStyle = {
  bg?: string;
  fc?: string;
  bl?: number;
  it?: number;
  un?: number;
};

// OOXML border style name → fortune-sheet numeric code.
// Mapping based on luckysheet's setBorder engine.
const BORDER_STYLE_CODES: Record<string, number> = {
  thin: 1,
  hair: 13,
  dotted: 4,
  dashed: 3,
  dashDot: 6,
  dashDotDot: 7,
  double: 8,
  medium: 9,
  mediumDashed: 10,
  mediumDashDot: 11,
  mediumDashDotDot: 12,
  slantDashDot: 5,
  thick: 2,
};

const mergeDxfInto = (
  eff: EffectiveStyle,
  dxf: Dxf | undefined,
  palette: ThemePalette,
) => {
  if (!dxf) return;
  if (dxf.fill?.bg) {
    const bg = resolveColor(dxf.fill.bg, palette);
    if (bg) eff.bg = bg;
  }
  if (dxf.font) {
    if (dxf.font.color) {
      const fc = resolveColor(dxf.font.color, palette);
      if (fc) eff.fc = fc;
    }
    if (dxf.font.bold !== undefined) eff.bl = dxf.font.bold ? 1 : 0;
    if (dxf.font.italic !== undefined) eff.it = dxf.font.italic ? 1 : 0;
    if (dxf.font.underline !== undefined) eff.un = dxf.font.underline ? 1 : 0;
  }
};

const mergeDxfBorderInto = (
  target: Partial<Record<BorderSide, { style: number; color: string }>>,
  dxf: Dxf | undefined,
  palette: ThemePalette,
) => {
  if (!dxf?.border) return;
  for (const side of Object.keys(dxf.border) as BorderSide[]) {
    const b = dxf.border[side];
    if (!b?.style) continue;
    const code = BORDER_STYLE_CODES[b.style];
    if (!code) continue;
    const color = resolveColor(b.color, palette) ?? '#000000';
    target[side] = { style: code, color };
  }
};

const computeEffectiveForCell = (
  table: TableDef,
  r: number,
  c: number,
  dxfs: Dxf[],
  elements: Partial<Record<TableElementType, number>>,
  palette: ThemePalette,
): EffectiveStyle => {
  const dxfFor = (name: TableElementType): Dxf | undefined => {
    const id = elements[name];
    return id !== undefined ? dxfs[id] : undefined;
  };
  const eff: EffectiveStyle = {};
  mergeDxfInto(eff, dxfFor('wholeTable'), palette);
  if (table.showHeaderRow && r === table.ref.r1) {
    mergeDxfInto(eff, dxfFor('headerRow'), palette);
  }
  if (table.showFirstCol && c === table.ref.c1) {
    mergeDxfInto(eff, dxfFor('firstColumn'), palette);
  }
  if (r === table.ref.r1 && c === table.ref.c1) {
    mergeDxfInto(eff, dxfFor('firstHeaderCell'), palette);
  }
  return eff;
};

const upsertCell = (
  celldata: NonNullable<Sheet['celldata']>,
  r: number,
  c: number,
): Cell => {
  const existing = celldata.find((entry) => entry.r === r && entry.c === c);
  if (existing) {
    if (!existing.v) existing.v = {};
    return existing.v as Cell;
  }
  const v: Cell = {};
  celldata.push({ r, c, v });
  return v;
};

const applyEffectiveToCell = (v: Cell, eff: EffectiveStyle) => {
  if (eff.bg !== undefined) v.bg = eff.bg;
  if (eff.fc !== undefined) v.fc = eff.fc;
  if (eff.bl !== undefined) v.bl = eff.bl;
  if (eff.it !== undefined) v.it = eff.it;
  if (eff.un !== undefined) v.un = eff.un;
};

type ResolvedBorder = { style: number; color: string };

const resolveHorizontal = (
  dxf: Dxf | undefined,
  palette: ThemePalette,
): ResolvedBorder | undefined => {
  const sides: Partial<Record<BorderSide, ResolvedBorder>> = {};
  mergeDxfBorderInto(sides, dxf, palette);
  return sides.horizontal;
};

const makeBorderEntry = (
  border: ResolvedBorder,
  r1: number,
  r2: number,
  c1: number,
  c2: number,
): object => ({
  rangeType: 'range',
  borderType: 'border-bottom',
  color: border.color,
  style: border.style,
  range: [{ row: [r1, r2], column: [c1, c2] }],
});

const emitBorderInfoForTable = (
  table: TableDef,
  dxfs: Dxf[],
  elements: Partial<Record<TableElementType, number>>,
  palette: ThemePalette,
): object[] => {
  const wholeDxf = elements.wholeTable !== undefined ? dxfs[elements.wholeTable] : undefined;
  const firstColDxf =
    table.showFirstCol && elements.firstColumn !== undefined
      ? dxfs[elements.firstColumn]
      : undefined;

  const wholeH = resolveHorizontal(wholeDxf, palette);
  const firstH = resolveHorizontal(firstColDxf, palette) ?? wholeH;

  if (!wholeH && !firstH) return [];

  const { r1, c1, r2, c2 } = table.ref;
  const entries: object[] = [];

  // Per-row border-bottom. r1 is excluded — that's the header row whose
  // bottom line is suppressed by headerRow.border.bottom in the sample.
  // r2 is also excluded because there's no line drawn below the last row
  // unless wholeTable specifies an outer bottom border.
  for (let r = r1 + 1; r <= r2 - 1; r++) {
    if (table.showFirstCol && firstH) {
      entries.push(makeBorderEntry(firstH, r, r, c1, c1));
    }
    if (wholeH) {
      const startCol = table.showFirstCol ? c1 + 1 : c1;
      if (startCol <= c2) {
        entries.push(makeBorderEntry(wholeH, r, r, startCol, c2));
      }
    }
  }

  return entries;
};

const applyTableToSheet = (
  sheet: Sheet,
  table: TableDef,
  tableStyles: Map<string, { elements: Partial<Record<TableElementType, number>> }>,
  dxfs: Dxf[],
  palette: ThemePalette,
) => {
  const style = tableStyles.get(table.styleName);
  if (!style) return;
  const { elements } = style;

  if (!sheet.celldata) sheet.celldata = [];
  const { r1, c1, r2, c2 }: CellRange = table.ref;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const eff = computeEffectiveForCell(table, r, c, dxfs, elements, palette);
      if (eff.bg === undefined && eff.fc === undefined &&
          eff.bl === undefined && eff.it === undefined && eff.un === undefined) {
        continue;
      }
      const v = upsertCell(sheet.celldata, r, c);
      applyEffectiveToCell(v, eff);
    }
  }

  if (!sheet.config) sheet.config = {};
  if (!sheet.config.borderInfo) sheet.config.borderInfo = [];
  const borders = emitBorderInfoForTable(table, dxfs, elements, palette);
  for (const entry of borders) sheet.config.borderInfo.push(entry);
};

export const applyTableStyles = async (
  file: File,
  sheets: Sheet[],
): Promise<Sheet[]> => {
  const zip = await loadXlsx(file);
  const themeDoc = zip.getXmlDoc('xl/theme/theme1.xml');
  const stylesDoc = zip.getXmlDoc('xl/styles.xml');
  const palette = parseTheme(themeDoc);
  const dxfs = parseDxfs(stylesDoc);
  const tableStyles = parseTableStyles(stylesDoc);
  const { tablesBySheetName } = indexSheetTables(zip.getXmlDoc, zip.listPaths);

  if (dxfs.length === 0 || tableStyles.size === 0 || tablesBySheetName.size === 0) {
    return sheets;
  }

  return sheets.map((sheet) => {
    const tableDocs = tablesBySheetName.get(sheet.name);
    if (!tableDocs || tableDocs.length === 0) return sheet;
    const next: Sheet = {
      ...sheet,
      celldata: sheet.celldata ? sheet.celldata.map((e) => ({ ...e, v: e.v ? { ...e.v } : e.v })) : [],
      config: {
        ...sheet.config,
        borderInfo: sheet.config?.borderInfo ? [...sheet.config.borderInfo] : [],
      },
    };
    for (const doc of tableDocs) {
      const table = parseTable(doc);
      if (!table) continue;
      applyTableToSheet(next, table, tableStyles, dxfs, palette);
    }
    return next;
  });
};
