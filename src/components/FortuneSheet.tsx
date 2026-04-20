import { useEffect, useRef, useState } from 'react';
import { Workbook } from '@fortune-sheet/react';
import type { WorkbookInstance } from '@fortune-sheet/react';
import type { Sheet } from '@fortune-sheet/core';
import '@fortune-sheet/react/dist/index.css';
import {
  FortuneExcelHelper,
  importToolBarItem,
  exportToolBarItem,
  transformExcelToFortune,
} from '@corbe30/fortune-excel';
import { applyTableStyles } from '../excel/applyTableStyles';
import { forceWrapOnOverflowCollisions } from '../excel/forceWrapOnOverflowCollisions';
import { backfillInlineStringV } from '../excel/backfillInlineStringV';

type Props = {
  file: File | null;
};

const emptySheets: Sheet[] = [{ name: 'Sheet1', celldata: [] }];

export const FortuneSheet = ({ file }: Props) => {
  const sheetRef = useRef<WorkbookInstance | null>(null);
  const [key, setKey] = useState(0);
  const [sheets, setSheets] = useState<Sheet[]>(emptySheets);

  useEffect(() => {
    if (!file) {
      setSheets(emptySheets);
      setKey((k) => k + 1);
      return;
    }
    let cancelled = false;
    let raw: Sheet[] | null = null;
    const capture = (next: Sheet[]) => {
      raw = next;
      setSheets(next);
    };
    // fortune-excel's transformExcelToFortune schedules a setTimeout that
    // calls sheetRef.current.setColumnWidth / setRowHeight for each sheet,
    // which races the Workbook's multi-sheet mount and throws
    // "sheet not found" on files with >1 sheet. fortune-sheet reads
    // config.columnlen / config.rowlen directly from the sheet data at
    // render time, so the imperative call is redundant. Pass a ref whose
    // `.current` is always null; fortune-excel uses optional-chaining, so
    // the entire imperative pass becomes a no-op.
    const shimRef = { current: null } as { current: null };
    transformExcelToFortune(file, capture, setKey, shimRef)
      .then(async () => {
        if (cancelled || !raw) return;
        const tablePatched = await applyTableStyles(file, raw);
        const inlineBackfilled = backfillInlineStringV(tablePatched);
        const patched = forceWrapOnOverflowCollisions(inlineBackfilled);
        if (!cancelled) {
          setSheets(patched);
          setKey((k) => k + 1);
        }
      })
      .catch((err) => {
        console.error('fortune-excel import failed', err);
      });
    return () => {
      cancelled = true;
    };
  }, [file]);

  return (
    <div className="flex h-full w-full flex-col">
      <FortuneExcelHelper
        setKey={setKey}
        setSheets={setSheets}
        sheetRef={sheetRef}
        config={{
          import: { xlsx: true, csv: true },
          export: { xlsx: true, csv: true },
        }}
      />
      <div className="min-h-[500px] flex-1 bg-white text-black">
        <Workbook
          key={key}
          ref={sheetRef}
          data={sheets}
          customToolbarItems={[importToolBarItem(), exportToolBarItem()]}
        />
      </div>
    </div>
  );
};
