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

const colorToCss = (
  color: { rgb?: string; theme?: number; tint?: number } | undefined,
): string | undefined => {
  if (!color?.rgb) return undefined;
  const rgb = color.rgb.length === 8 ? color.rgb.slice(2) : color.rgb;
  return `#${rgb}`;
};

type XlsxColor = { rgb?: string; theme?: number; tint?: number };
type XlsxFill =
  | { type: 'solid'; color?: XlsxColor }
  | { type: 'pattern'; foreground?: XlsxColor; background?: XlsxColor }
  | { type: 'gradient' };

type StyleMap = {
  bold: boolean | undefined;
  italic: boolean | undefined;
  underline: boolean | undefined;
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

  const style: React.CSSProperties = {};

  if (cell.style('bold')) style.fontWeight = 'bold';
  if (cell.style('italic')) style.fontStyle = 'italic';
  if (cell.style('underline')) style.textDecoration = 'underline';

  const fontSize = cell.style('fontSize');
  if (fontSize) style.fontSize = `${fontSize}px`;

  const fontFamily = cell.style('fontFamily');
  if (fontFamily) style.fontFamily = fontFamily;

  const fontColor = colorToCss(cell.style('fontColor'));
  if (fontColor) style.color = fontColor;

  const fill = cell.style('fill');
  if (fill?.type === 'solid') {
    const bg = colorToCss(fill.color);
    if (bg) style.backgroundColor = bg;
  } else if (fill?.type === 'pattern') {
    const bg = colorToCss(fill.foreground ?? fill.background);
    if (bg) style.backgroundColor = bg;
  }

  const horizontal = cell.style('horizontalAlignment');
  if (horizontal) style.textAlign = horizontal as React.CSSProperties['textAlign'];

  const vertical = cell.style('verticalAlignment');
  if (vertical === 'center') style.verticalAlign = 'middle';
  else if (vertical) style.verticalAlign = vertical as React.CSSProperties['verticalAlign'];

  return { value, style };
};

const StyledViewer: DataViewerComponent<StyledCell> = ({
  cell,
}: DataViewerProps<StyledCell>) => {
  if (!cell) return null;
  return (
    <div
      style={{
        width: '100%',
        height: '100%',
        padding: '4px',
        boxSizing: 'border-box',
        ...cell.style,
      }}
    >
      {cell.value}
    </div>
  );
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
