import { useState } from 'react';
import { FortuneSheet } from './components/FortuneSheet';

export const App = () => {
  const [file, setFile] = useState<File | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const next = e.target.files?.[0];
    setFile(next ?? null);
  };

  return (
    <div className="flex min-h-screen flex-col bg-gray-900 p-6 text-white">
      <header className="mb-4 flex flex-col gap-2">
        <h1 className="text-xl font-semibold">XLSX read-only viewer</h1>
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

      <div className="flex flex-col items-center gap-4">
        <section className="flex w-fit max-w-full flex-col rounded border border-gray-700">
          <div>
            <FortuneSheet file={file} />
          </div>
        </section>
      </div>
    </div>
  );
};
