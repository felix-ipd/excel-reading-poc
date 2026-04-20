export type TableElementType =
  | 'wholeTable'
  | 'headerRow'
  | 'firstColumn'
  | 'firstColumnStripe'
  | 'firstHeaderCell'
  | 'lastColumn'
  | 'totalRow'
  | 'firstRowStripe'
  | 'secondRowStripe'
  | 'firstColumnStripe'
  | 'secondColumnStripe';

export type TableStyle = {
  elements: Partial<Record<TableElementType, number>>;
};

export const parseTableStyles = (
  stylesDoc: Document | null,
): Map<string, TableStyle> => {
  const out = new Map<string, TableStyle>();
  if (!stylesDoc) return out;
  const styles = stylesDoc.getElementsByTagNameNS('*', 'tableStyles')[0];
  if (!styles) return out;

  for (let i = 0; i < styles.children.length; i++) {
    const styleEl = styles.children[i];
    if (styleEl.localName !== 'tableStyle') continue;
    const name = styleEl.getAttribute('name');
    if (!name) continue;
    const elements: TableStyle['elements'] = {};
    for (let j = 0; j < styleEl.children.length; j++) {
      const el = styleEl.children[j];
      if (el.localName !== 'tableStyleElement') continue;
      const type = el.getAttribute('type');
      const dxfId = el.getAttribute('dxfId');
      if (!type || dxfId === null) continue;
      elements[type as TableElementType] = Number(dxfId);
    }
    out.set(name, { elements });
  }
  return out;
};
