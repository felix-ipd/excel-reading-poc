import { useState } from 'react';
import { FortuneSheet } from './components/FortuneSheet';
import { PopulateAndSpreadsheet } from './components/PopulateAndSpreadsheet';

export const App = () => {
  const [file, setFile] = useState<File | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const next = e.target.files?.[0];
    setFile(next ?? null);
  };

  return (
    <div className="flex min-h-screen flex-col bg-gray-900 p-6 text-white">
      <header className="mb-4 flex flex-col gap-2">
        <h1 className="text-xl font-semibold">
          XLSX rendering comparison: fortune-sheet vs xlsx-populate + react-spreadsheet
        </h1>
        <label className="flex items-center gap-2 text-sm">
          <span>Select an .xlsx file:</span>
          <input
            type="file"
            accept=".xlsx,.csv"
            onChange={handleFileChange}
            className="text-white"
          />
          {file && (
            <span className="text-gray-400">
              Loaded: <strong>{file.name}</strong>
            </span>
          )}
        </label>
      </header>

      <div className="grid flex-1 grid-cols-1 gap-4 lg:grid-cols-2">
        <section className="flex flex-col rounded border border-gray-700">
          <h2 className="bg-gray-800 px-3 py-2 text-sm font-semibold">
            fortune-sheet + fortune-excel
          </h2>
          <div className="flex-1">
            <FortuneSheet file={file} />
          </div>
        </section>

        <section className="flex flex-col rounded border border-gray-700">
          <h2 className="bg-gray-800 px-3 py-2 text-sm font-semibold">
            xlsx-populate + react-spreadsheet
          </h2>
          <div className="flex-1">
            <PopulateAndSpreadsheet file={file} />
          </div>
        </section>
      </div>
    </div>
  );
};
