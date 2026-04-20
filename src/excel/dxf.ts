import { readColorRef, type ColorRef } from './colors';

export type BorderSide = 'horizontal' | 'vertical' | 'top' | 'bottom' | 'left' | 'right';

export type DxfBorder = { style?: string; color?: ColorRef };

export type Dxf = {
  fill?: { bg?: ColorRef };
  font?: {
    color?: ColorRef;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  };
  border?: Partial<Record<BorderSide, DxfBorder>>;
};

const firstChildByTag = (parent: Element, local: string): Element | null => {
  for (let i = 0; i < parent.children.length; i++) {
    if (parent.children[i].localName === local) return parent.children[i];
  }
  return null;
};

const readFill = (fillEl: Element): Dxf['fill'] => {
  const pattern = firstChildByTag(fillEl, 'patternFill');
  if (!pattern) return undefined;
  const bgEl = firstChildByTag(pattern, 'bgColor') ?? firstChildByTag(pattern, 'fgColor');
  const bg = readColorRef(bgEl);
  return bg ? { bg } : undefined;
};

const readFont = (fontEl: Element): Dxf['font'] => {
  const font: NonNullable<Dxf['font']> = {};
  for (let i = 0; i < fontEl.children.length; i++) {
    const child = fontEl.children[i];
    switch (child.localName) {
      case 'b':
        font.bold = child.getAttribute('val') !== '0';
        break;
      case 'i':
        font.italic = child.getAttribute('val') !== '0';
        break;
      case 'u':
        font.underline = child.getAttribute('val') !== 'none';
        break;
      case 'color': {
        const c = readColorRef(child);
        if (c) font.color = c;
        break;
      }
      default:
        break;
    }
  }
  return Object.keys(font).length > 0 ? font : undefined;
};

const readBorder = (borderEl: Element): Dxf['border'] => {
  const result: Partial<Record<BorderSide, DxfBorder>> = {};
  const sides: BorderSide[] = ['horizontal', 'vertical', 'top', 'bottom', 'left', 'right'];
  for (const side of sides) {
    const sideEl = firstChildByTag(borderEl, side);
    if (!sideEl) continue;
    const style = sideEl.getAttribute('style') ?? undefined;
    const colorEl = firstChildByTag(sideEl, 'color');
    const color = readColorRef(colorEl);
    if (style || color) result[side] = { style, color };
  }
  return Object.keys(result).length > 0 ? result : undefined;
};

export const parseDxfs = (stylesDoc: Document | null): Dxf[] => {
  if (!stylesDoc) return [];
  const dxfsEl = stylesDoc.getElementsByTagNameNS('*', 'dxfs')[0];
  if (!dxfsEl) return [];
  const out: Dxf[] = [];
  for (let i = 0; i < dxfsEl.children.length; i++) {
    const dxfEl = dxfsEl.children[i];
    if (dxfEl.localName !== 'dxf') {
      out.push({});
      continue;
    }
    const dxf: Dxf = {};
    const fontEl = firstChildByTag(dxfEl, 'font');
    const fillEl = firstChildByTag(dxfEl, 'fill');
    const borderEl = firstChildByTag(dxfEl, 'border');
    if (fontEl) dxf.font = readFont(fontEl);
    if (fillEl) dxf.fill = readFill(fillEl);
    if (borderEl) dxf.border = readBorder(borderEl);
    out.push(dxf);
  }
  return out;
};
