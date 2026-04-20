export type CellRange = { r1: number; c1: number; r2: number; c2: number };

export type TableDef = {
  name: string;
  ref: CellRange;
  styleName: string;
  showHeaderRow: boolean;
  showFirstCol: boolean;
  showLastCol: boolean;
  showRowStripes: boolean;
  showColStripes: boolean;
};

// Convert a spreadsheet column label (e.g. "AB") to a 0-based index.
const colFromLabel = (label: string): number => {
  let n = 0;
  for (const ch of label) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
};

export const parseRef = (ref: string): CellRange | null => {
  const parts = ref.split(':');
  const m1 = parts[0].match(/^([A-Z]+)([0-9]+)$/);
  if (!m1) return null;
  const c1 = colFromLabel(m1[1]);
  const r1 = Number(m1[2]) - 1;
  if (parts.length === 1) {
    return { r1, c1, r2: r1, c2: c1 };
  }
  const m2 = parts[1].match(/^([A-Z]+)([0-9]+)$/);
  if (!m2) return null;
  const c2 = colFromLabel(m2[1]);
  const r2 = Number(m2[2]) - 1;
  return { r1, c1, r2, c2 };
};

export const parseTable = (doc: Document): TableDef | null => {
  const tableEl = doc.documentElement;
  if (!tableEl || tableEl.localName !== 'table') return null;

  const name = tableEl.getAttribute('name') ?? tableEl.getAttribute('displayName') ?? '';
  const ref = parseRef(tableEl.getAttribute('ref') ?? '');
  if (!ref) return null;

  const styleInfo = tableEl.getElementsByTagNameNS('*', 'tableStyleInfo')[0];
  const styleName = styleInfo?.getAttribute('name') ?? '';
  const headerRowCount = Number(tableEl.getAttribute('headerRowCount') ?? '1');

  return {
    name,
    ref,
    styleName,
    showHeaderRow: headerRowCount > 0,
    showFirstCol: styleInfo?.getAttribute('showFirstColumn') === '1',
    showLastCol: styleInfo?.getAttribute('showLastColumn') === '1',
    showRowStripes: styleInfo?.getAttribute('showRowStripes') === '1',
    showColStripes: styleInfo?.getAttribute('showColumnStripes') === '1',
  };
};

// Given the workbook rels graph, resolve which tables belong to which sheet.
export type SheetTableIndex = {
  sheetNameByPath: Map<string, string>;
  tablesBySheetName: Map<string, Document[]>;
};

const resolvePath = (basePath: string, target: string): string => {
  // target may be relative, e.g. "../tables/table1.xml" from "xl/worksheets/_rels/sheet1.xml.rels"
  const stack = basePath.split('/').slice(0, -1);
  for (const part of target.split('/')) {
    if (part === '..') stack.pop();
    else if (part !== '.') stack.push(part);
  }
  return stack.join('/');
};

export const indexSheetTables = (
  getXmlDoc: (path: string) => Document | null,
  listPaths: () => string[],
): SheetTableIndex => {
  const sheetNameByPath = new Map<string, string>();
  const tablesBySheetName = new Map<string, Document[]>();

  const workbookDoc = getXmlDoc('xl/workbook.xml');
  const workbookRels = getXmlDoc('xl/_rels/workbook.xml.rels');
  if (!workbookDoc || !workbookRels) return { sheetNameByPath, tablesBySheetName };

  const relTargetById = new Map<string, string>();
  const rels = workbookRels.getElementsByTagNameNS('*', 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    const id = rels[i].getAttribute('Id');
    const target = rels[i].getAttribute('Target');
    if (id && target) relTargetById.set(id, target);
  }

  const sheetEls = workbookDoc.getElementsByTagNameNS('*', 'sheet');
  for (let i = 0; i < sheetEls.length; i++) {
    const sh = sheetEls[i];
    const name = sh.getAttribute('name') ?? '';
    // r:id on <sheet> sits in the relationships namespace. Try a few lookups
    // because browsers disagree on whether the prefixed form is accessible.
    const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const rid =
      sh.getAttributeNS(relNs, 'id') ||
      sh.getAttribute('r:id') ||
      sh.getAttribute('id') ||
      '';
    const target = rid ? relTargetById.get(rid) : undefined;
    if (!target) continue;
    const sheetPath = target.startsWith('/')
      ? target.slice(1)
      : resolvePath('xl/workbook.xml', target);
    sheetNameByPath.set(sheetPath, name);

    // Look up per-sheet rels for <table> entries
    const sheetRelPath = `${sheetPath.split('/').slice(0, -1).join('/')}/_rels/${
      sheetPath.split('/').pop() ?? ''
    }.rels`;
    const sheetRels = getXmlDoc(sheetRelPath);
    if (!sheetRels) continue;
    const relEls = sheetRels.getElementsByTagNameNS('*', 'Relationship');
    const docs: Document[] = [];
    for (let j = 0; j < relEls.length; j++) {
      const rel = relEls[j];
      const type = rel.getAttribute('Type') ?? '';
      if (!type.endsWith('/table')) continue;
      const relTarget = rel.getAttribute('Target') ?? '';
      const tablePath = relTarget.startsWith('/')
        ? relTarget.slice(1)
        : resolvePath(sheetPath, relTarget);
      const tableDoc = getXmlDoc(tablePath);
      if (tableDoc) docs.push(tableDoc);
    }
    if (docs.length > 0) tablesBySheetName.set(name, docs);
  }

  // Reference listPaths to avoid unused-arg lint in case the surface broadens.
  void listPaths;

  return { sheetNameByPath, tablesBySheetName };
};
