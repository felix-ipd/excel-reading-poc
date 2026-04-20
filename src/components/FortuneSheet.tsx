import { useEffect, useMemo, useReducer, useRef } from 'react';
import { Workbook } from '@fortune-sheet/react';
import type { WorkbookInstance } from '@fortune-sheet/react';
import type { Sheet } from '@fortune-sheet/core';
import '@fortune-sheet/react/dist/index.css';
import { transformExcelToFortune } from '@corbe30/fortune-excel';
import { applyTableStyles } from '../excel/applyTableStyles';
import { forceWrapOnOverflowCollisions } from '../excel/forceWrapOnOverflowCollisions';
import { backfillInlineStringV } from '../excel/backfillInlineStringV';
import { bindDisplayBounds } from '../excel/bindDisplayBounds';
import { measureSheet } from '../excel/measureSheet';

type Props = {
  file: File | null;
};

const emptySheets: Sheet[] = [{ name: 'Sheet1', celldata: [] }];

type ViewerState = {
  sheets: Sheet[];
  activeSheetIndex: number;
  key: number;
};

type ViewerAction =
  | { type: 'reset' }
  | { type: 'load'; sheets: Sheet[] }
  | { type: 'selectSheet'; index: number };

const initialState: ViewerState = {
  sheets: emptySheets,
  activeSheetIndex: 0,
  key: 0,
};

const viewerReducer = (state: ViewerState, action: ViewerAction): ViewerState => {
  switch (action.type) {
    case 'reset':
      return { sheets: emptySheets, activeSheetIndex: 0, key: state.key + 1 };
    case 'load':
      return { sheets: action.sheets, activeSheetIndex: 0, key: state.key + 1 };
    case 'selectSheet':
      return { ...state, activeSheetIndex: action.index, key: state.key + 1 };
  }
};

export const FortuneSheet = ({ file }: Props) => {
  const sheetRef = useRef<WorkbookInstance | null>(null);
  const [state, dispatch] = useReducer(viewerReducer, initialState);
  const { sheets, activeSheetIndex, key } = state;

  useEffect(() => {
    if (!file) {
      dispatch({ type: 'reset' });
      return;
    }
    let cancelled = false;
    let raw: Sheet[] | null = null;
    const capture = (next: Sheet[]) => {
      raw = next;
    };
    // fortune-excel's transformExcelToFortune schedules a setTimeout that
    // calls sheetRef.current.setColumnWidth / setRowHeight for each sheet,
    // which races the Workbook's multi-sheet mount and throws
    // "sheet not found" on files with >1 sheet. fortune-sheet reads
    // config.columnlen / config.rowlen directly from the sheet data at
    // render time, so the imperative call is redundant. Pass a ref whose
    // `.current` is always null; fortune-excel uses optional-chaining, so
    // the entire imperative pass becomes a no-op.
    const shimRef = { current: null };
    // No-op setter: fortune-excel's internal setKey call after `capture`
    // would otherwise force an extra Workbook remount with the raw
    // (unpatched) data right before our post-processors replace it.
    const noop = () => {};
    transformExcelToFortune(file, capture, noop, shimRef)
      .then(async () => {
        if (cancelled || !raw) return;
        const tablePatched = await applyTableStyles(file, raw);
        const inlineBackfilled = backfillInlineStringV(tablePatched);
        const wrapPatched = forceWrapOnOverflowCollisions(inlineBackfilled);
        const bounded = bindDisplayBounds(wrapPatched);
        if (!cancelled) {
          dispatch({ type: 'load', sheets: bounded });
        }
      })
      .catch((err) => {
        console.error('fortune-excel import failed', err);
      });
    return () => {
      cancelled = true;
    };
  }, [file]);

  const safeIndex = Math.min(activeSheetIndex, Math.max(0, sheets.length - 1));
  const activeSheet = sheets[safeIndex] ?? emptySheets[0];
  const { width, height } = useMemo(() => measureSheet(activeSheet), [activeSheet]);

  return (
    <div className="flex flex-col gap-2">
      {/* {sheets.length > 1 && ( */}
      {/*   <select */}
      {/*     value={safeIndex} */}
      {/*     onChange={(e) => dispatch({ type: 'selectSheet', index: Number(e.target.value) })} */}
      {/*     className="self-start rounded border border-gray-300 bg-white px-3 py-1 text-sm text-black" */}
      {/*   > */}
      {/*     {sheets.map((s, i) => ( */}
      {/*       <option key={s.id ?? `${i}-${s.name}`} value={i}> */}
      {/*         {s.name} */}
      {/*       </option> */}
      {/*     ))} */}
      {/*   </select> */}
      {/* )} */}
      <div
        className="overflow-auto"
        style={{
          maxWidth: 'calc(100vw - 4rem)',
          maxHeight: 'calc(100vh - 12rem)',
        }}
      >
        <div
          className="bg-white text-black"
          style={{ width, height }}
        >
          <Workbook
            key={`${key}-${safeIndex}`}
            ref={sheetRef}
            data={[activeSheet]}
            allowEdit={false}
            showToolbar={false}
            showFormulaBar={false}
            showSheetTabs={false}
            cellContextMenu={[]}
            headerContextMenu={[]}
          />
        </div>
      </div>
    </div>
  );
};
