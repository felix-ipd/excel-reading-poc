export type ColorRef = {
  rgb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
  auto?: boolean;
};

export type ThemePalette = string[];

// OOXML legacy indexed palette (62 entries). Rarely used in modern files,
// but we include the first ~16 for safety and fall through to undefined for
// unusual indexes.
const INDEXED_PALETTE = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
];

// HLS tint per OOXML 18.8.19.
export const applyTint = (hex: string, tint: number | undefined): string => {
  if (!tint) return hex;
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  const l0 = (max + min) / 2;
  const d = max - min;
  let s = 0;
  if (d !== 0) {
    s = l0 > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      default:
        h = (r - g) / d + 4;
    }
    h /= 6;
  }

  const t = Math.max(-1, Math.min(1, tint));
  const l = t < 0 ? l0 * (1 + t) : l0 * (1 - t) + t;

  const hue2rgb = (p: number, q: number, tt: number) => {
    let v = tt;
    if (v < 0) v += 1;
    if (v > 1) v -= 1;
    if (v < 1 / 6) return p + (q - p) * 6 * v;
    if (v < 1 / 2) return q;
    if (v < 2 / 3) return p + (q - p) * (2 / 3 - v) * 6;
    return p;
  };

  let r2: number;
  let g2: number;
  let b2: number;
  if (s === 0) {
    r2 = g2 = b2 = l;
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r2 = hue2rgb(p, q, h + 1 / 3);
    g2 = hue2rgb(p, q, h);
    b2 = hue2rgb(p, q, h - 1 / 3);
  }

  const toHex = (v: number) =>
    Math.max(0, Math.min(255, Math.round(v * 255))).toString(16).padStart(2, '0');
  return `${toHex(r2)}${toHex(g2)}${toHex(b2)}`;
};

// SpreadsheetML theme indexes swap the first two pairs vs. clrScheme XML order.
const remapThemeIndex = (idx: number): number => {
  if (idx === 0) return 1;
  if (idx === 1) return 0;
  if (idx === 2) return 3;
  if (idx === 3) return 2;
  return idx;
};

export const resolveColor = (
  ref: ColorRef | undefined,
  palette: ThemePalette,
): string | undefined => {
  if (!ref) return undefined;
  if (ref.auto) return undefined;

  let hex: string | undefined;
  if (ref.rgb) {
    hex = ref.rgb.length === 8 ? ref.rgb.slice(2) : ref.rgb;
  } else if (typeof ref.theme === 'number') {
    hex = palette[remapThemeIndex(ref.theme)];
  } else if (typeof ref.indexed === 'number') {
    hex = INDEXED_PALETTE[ref.indexed];
  }
  if (!hex) return undefined;
  return `#${applyTint(hex, ref.tint).toUpperCase()}`;
};

export const readColorRef = (el: Element | null): ColorRef | undefined => {
  if (!el) return undefined;
  const ref: ColorRef = {};
  const rgb = el.getAttribute('rgb');
  const theme = el.getAttribute('theme');
  const tint = el.getAttribute('tint');
  const indexed = el.getAttribute('indexed');
  const auto = el.getAttribute('auto');
  if (rgb) ref.rgb = rgb;
  if (theme !== null) ref.theme = Number(theme);
  if (tint !== null) ref.tint = Number(tint);
  if (indexed !== null) ref.indexed = Number(indexed);
  if (auto === '1') ref.auto = true;
  return Object.keys(ref).length > 0 ? ref : undefined;
};
