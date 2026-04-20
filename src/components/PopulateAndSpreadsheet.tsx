import React, { useEffect, useMemo, useState } from 'react';
import Spreadsheet from 'react-spreadsheet';
import type { DataViewerComponent, DataViewerProps } from 'react-spreadsheet';
// @ts-expect-error no types shipped for the browser build
import XlsxPopulate from 'xlsx-populate/browser/xlsx-populate';

type StyledCell = {
  value: string | number;
  style?: React.CSSProperties;
};

type Props = {
  file: File | null;
};

type XlsxColor = { rgb?: string; theme?: number; tint?: number };
type XlsxFill =
  | { type: 'solid'; color?: XlsxColor }
  | { type: 'pattern'; foreground?: XlsxColor; background?: XlsxColor }
  | { type: 'gradient' };

// Office 2013+ default theme palette. Used as a fallback when the workbook
// references theme colors but we don't parse the per-file theme1.xml.
// Index order follows SpreadsheetML's `theme=""` attribute (not DrawingML's
// internal ordering), so 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6.
const THEME_FALLBACK = [
  'FFFFFF', // lt1 (background 1)
  '000000', // dk1 (text 1)
  'E7E6E6', // lt2 (background 2)
  '44546A', // dk2 (text 2)
  '4472C4', // accent1
  'ED7D31', // accent2
  'A5A5A5', // accent3
  'FFC000', // accent4
  '5B9BD5', // accent5
  '70AD47', // accent6
  '0563C1', // hyperlink
  '954F72', // followed hyperlink
];

const hexToRgb = (hex: string): [number, number, number] => {
  const h = hex.length === 8 ? hex.slice(2) : hex;
  return [
    parseInt(h.slice(0, 2), 16),
    parseInt(h.slice(2, 4), 16),
    parseInt(h.slice(4, 6), 16),
  ];
};

const rgbToHex = (r: number, g: number, b: number): string => {
  const clamp = (n: number) => Math.max(0, Math.min(255, Math.round(n)));
  return [r, g, b].map((n) => clamp(n).toString(16).padStart(2, '0')).join('');
};

// OOXML tint: [-1, 1]. Applied in HSL luminance space. Cheap RGB approximation
// is close enough for PoC fidelity.
const applyTint = (hex: string, tint: number | undefined): string => {
  if (!tint) return hex;
  const [r, g, b] = hexToRgb(hex);
  const t = Math.max(-1, Math.min(1, tint));
  const shift = (c: number) => (t < 0 ? c * (1 + t) : c + (255 - c) * t);
  return rgbToHex(shift(r), shift(g), shift(b));
};

const resolveColor = (color: XlsxColor | undefined): string | undefined => {
  if (!color) return undefined;
  let hex: string | undefined;
  if (color.rgb) {
    hex = color.rgb.length === 8 ? color.rgb.slice(2) : color.rgb;
  } else if (typeof color.theme === 'number') {
    hex = THEME_FALLBACK[color.theme];
  }
  if (!hex) return undefined;
  return `#${applyTint(hex, color.tint)}`;
};

type StyleMap = {
  bold: boolean | undefined;
  italic: boolean | undefined;
  underline: boolean | undefined;
  strikethrough: boolean | undefined;
  fontSize: number | undefined;
  fontFamily: string | undefined;
  fontColor: XlsxColor | undefined;
  fill: XlsxFill | undefined;
  horizontalAlignment: string | undefined;
  verticalAlignment: string | undefined;
};

type XlsxCell = {
  value: () => unknown;
  style: <K extends keyof StyleMap>(name: K) => StyleMap[K];
};

const verticalToFlex = (v: string | undefined): string => {
  switch (v) {
    case 'top':
      return 'flex-start';
    case 'center':
      return 'center';
    case 'bottom':
    default:
      return 'flex-end';
  }
};

const horizontalToFlex = (h: string | undefined): string => {
  switch (h) {
    case 'center':
    case 'centerContinuous':
      return 'center';
    case 'right':
    case 'end':
      return 'flex-end';
    case 'left':
    case 'start':
    default:
      return 'flex-start';
  }
};

const toStyledCell = (cell: XlsxCell): StyledCell => {
  const rawValue = cell.value();
  let value: string | number = '';
  if (rawValue && typeof rawValue === 'object' && 'text' in rawValue) {
    const text = (rawValue as { text?: () => string }).text;
    value = text?.() ?? '';
  } else if (typeof rawValue === 'string' || typeof rawValue === 'number') {
    value = rawValue;
  } else if (rawValue != null) {
    value = String(rawValue);
  }

  const style: React.CSSProperties = {
    display: 'flex',
    width: '100%',
    height: '100%',
    padding: '2px 4px',
    boxSizing: 'border-box',
    lineHeight: 1.2,
    overflow: 'hidden',
  };

  if (cell.style('bold')) style.fontWeight = 'bold';
  if (cell.style('italic')) style.fontStyle = 'italic';

  const decorations: string[] = [];
  if (cell.style('underline')) decorations.push('underline');
  if (cell.style('strikethrough')) decorations.push('line-through');
  if (decorations.length > 0) style.textDecoration = decorations.join(' ');

  const fontSize = cell.style('fontSize');
  if (fontSize) style.fontSize = `${fontSize}px`;

  const fontFamily = cell.style('fontFamily');
  if (fontFamily) style.fontFamily = fontFamily;

  const fontColor = resolveColor(cell.style('fontColor'));
  if (fontColor) style.color = fontColor;

  const fill = cell.style('fill');
  if (fill?.type === 'solid') {
    const bg = resolveColor(fill.color);
    if (bg) style.backgroundColor = bg;
  } else if (fill?.type === 'pattern') {
    const bg = resolveColor(fill.foreground ?? fill.background);
    if (bg) style.backgroundColor = bg;
  }

  const horizontal = cell.style('horizontalAlignment');
  style.justifyContent = horizontalToFlex(horizontal);
  if (horizontal) {
    style.textAlign = horizontal as React.CSSProperties['textAlign'];
  }

  const vertical = cell.style('verticalAlignment');
  style.alignItems = verticalToFlex(vertical);

  return { value, style };
};

const StyledViewer: DataViewerComponent<StyledCell> = ({
  cell,
}: DataViewerProps<StyledCell>) => {
  if (!cell) return null;
  return <div style={cell.style}>{cell.value}</div>;
};

export const PopulateAndSpreadsheet = ({ file }: Props) => {
  const [data, setData] = useState<StyledCell[][]>([]);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!file) {
      setData([]);
      setError(null);
      return;
    }

    let cancelled = false;

    const run = async () => {
      try {
        setError(null);
        const arrayBuffer = await file.arrayBuffer();
        const workbook = await XlsxPopulate.fromDataAsync(arrayBuffer);
        const sheet = workbook.sheet(0);
        const usedRange = sheet.usedRange();

        if (!usedRange) {
          if (!cancelled) setData([]);
          return;
        }

        const endCell = usedRange.endCell();
        const rowCount = endCell.rowNumber();
        const colCount = endCell.columnNumber();

        const sheetData: StyledCell[][] = [];
        for (let r = 1; r <= rowCount; r++) {
          const row: StyledCell[] = [];
          for (let c = 1; c <= colCount; c++) {
            row.push(toStyledCell(sheet.cell(r, c)));
          }
          sheetData.push(row);
        }

        if (!cancelled) setData(sheetData);
      } catch (err) {
        console.error('xlsx-populate import failed', err);
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      }
    };

    run();
    return () => {
      cancelled = true;
    };
  }, [file]);

  const tableData = useMemo(() => data, [data]);

  return (
    <div className="flex h-full w-full flex-col bg-white text-black">
      {error && <div className="p-2 text-red-600">{error}</div>}
      <div className="min-h-[500px] flex-1 overflow-auto p-2">
        <Spreadsheet data={tableData} DataViewer={StyledViewer} />
      </div>
    </div>
  );
};
