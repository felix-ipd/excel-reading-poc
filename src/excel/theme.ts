import type { ThemePalette } from './colors';

// Office 2013+ default palette in clrScheme XML order
// (dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink).
const DEFAULT_PALETTE: ThemePalette = [
  '000000', 'FFFFFF', '44546A', 'E7E6E6',
  '4472C4', 'ED7D31', 'A5A5A5', 'FFC000',
  '5B9BD5', '70AD47', '0563C1', '954F72',
];

const SCHEME_ORDER = [
  'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4',
  'accent5', 'accent6', 'hlink', 'folHlink',
];

// Theme1.xml uses an "a:" namespace. We match by localName to stay namespace-agnostic.
const findChildByLocalName = (parent: Element, local: string): Element | null => {
  for (let i = 0; i < parent.children.length; i++) {
    const c = parent.children[i];
    if (c.localName === local) return c;
  }
  return null;
};

const readSchemeColor = (slot: Element): string | undefined => {
  const srgb = findChildByLocalName(slot, 'srgbClr');
  if (srgb) {
    const val = srgb.getAttribute('val');
    if (val) return val.toUpperCase();
  }
  const sys = findChildByLocalName(slot, 'sysClr');
  if (sys) {
    const last = sys.getAttribute('lastClr');
    if (last) return last.toUpperCase();
  }
  return undefined;
};

export const parseTheme = (doc: Document | null): ThemePalette => {
  if (!doc) return [...DEFAULT_PALETTE];

  const clrScheme = doc.getElementsByTagNameNS('*', 'clrScheme')[0];
  if (!clrScheme) return [...DEFAULT_PALETTE];

  const palette: ThemePalette = [];
  for (const name of SCHEME_ORDER) {
    const slot = findChildByLocalName(clrScheme, name);
    const hex = slot ? readSchemeColor(slot) : undefined;
    palette.push(hex ?? DEFAULT_PALETTE[palette.length]);
  }
  return palette;
};
