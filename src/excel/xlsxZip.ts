import JSZip from 'jszip';

export type XlsxZip = {
  getXmlDoc: (path: string) => Document | null;
  getText: (path: string) => string | null;
  listPaths: () => string[];
};

const parser = new DOMParser();

export const loadXlsx = async (file: File): Promise<XlsxZip> => {
  const zip = await JSZip.loadAsync(file);
  const texts = new Map<string, string>();
  await Promise.all(
    Object.keys(zip.files).map(async (path) => {
      const entry = zip.files[path];
      if (entry.dir) return;
      const content = await entry.async('string');
      texts.set(path, content);
    }),
  );

  return {
    getXmlDoc: (path) => {
      const text = texts.get(path);
      if (!text) return null;
      const doc = parser.parseFromString(text, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length > 0) return null;
      return doc;
    },
    getText: (path) => texts.get(path) ?? null,
    listPaths: () => Array.from(texts.keys()),
  };
};
