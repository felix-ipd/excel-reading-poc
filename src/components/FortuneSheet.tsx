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
    transformExcelToFortune(file, setSheets, setKey, sheetRef).catch((err) => {
      console.error('fortune-excel import failed', err);
    });
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
