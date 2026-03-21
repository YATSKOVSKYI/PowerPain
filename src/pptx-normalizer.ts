import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

export interface NormalizeOptions {
  targetFont: string;
  fixLanguage: boolean;
  fixTheme: boolean;
  fixSlides: boolean;
  fixLayouts: boolean;
  fixMaster: boolean;
}

export interface NormalizeStats {
  slides: number;
  layouts: number;
  masters: number;
  themes: number;
}

const DEFAULT_OPTIONS: NormalizeOptions = {
  targetFont: "Arial",
  fixLanguage: true,
  fixTheme: true,
  fixSlides: true,
  fixLayouts: true,
  fixMaster: true,
};

// ─── XML font normalization ─────────────────────────────────────────────────

function normalizeXmlFonts(xmlString: string, opts: NormalizeOptions): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "text/xml");

  patchRunProperties(doc, "rPr", opts);
  patchRunProperties(doc, "defRPr", opts);
  patchRunProperties(doc, "endParaRPr", opts);

  let result = new XMLSerializer().serializeToString(doc);

  // Replace remaining theme references via regex
  result = result.replace(/typeface="\+mj-lt"/g, `typeface="${opts.targetFont}"`);
  result = result.replace(/typeface="\+mn-lt"/g, `typeface="${opts.targetFont}"`);
  result = result.replace(/typeface="\+mj-ea"/g, `typeface="${opts.targetFont}"`);
  result = result.replace(/typeface="\+mn-ea"/g, `typeface="${opts.targetFont}"`);

  return result;
}

function patchRunProperties(
  doc: Document,
  tagName: string,
  opts: NormalizeOptions
): void {
  const elements = doc.getElementsByTagNameNS(NS_A, tagName);

  for (let i = 0; i < elements.length; i++) {
    const rPr = elements[i];

    // Fix CJK language → en-US so PowerPoint uses <a:latin> instead of <a:ea>
    if (opts.fixLanguage) {
      const lang = rPr.getAttribute("lang");
      if (lang && isCJKLocale(lang)) {
        rPr.setAttribute("lang", "en-US");
        rPr.removeAttribute("altLang");
      }
    }

    removeChildrenByLocalName(rPr, ["latin", "ea", "cs", "sym"]);

    const latin = doc.createElementNS(NS_A, "a:latin");
    latin.setAttribute("typeface", opts.targetFont);
    rPr.appendChild(latin);

    const ea = doc.createElementNS(NS_A, "a:ea");
    ea.setAttribute("typeface", opts.targetFont);
    rPr.appendChild(ea);
  }
}

// ─── Theme font normalization ───────────────────────────────────────────────

function normalizeThemeFonts(xmlString: string, targetFont: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "text/xml");

  for (const fontGroup of ["majorFont", "minorFont"] as const) {
    const group = doc.getElementsByTagNameNS(NS_A, fontGroup)[0];
    if (!group) continue;

    let latin = group.getElementsByTagNameNS(NS_A, "latin")[0];
    if (!latin) {
      latin = doc.createElementNS(NS_A, "a:latin");
      group.insertBefore(latin, group.firstChild);
    }
    latin.setAttribute("typeface", targetFont);
  }

  return new XMLSerializer().serializeToString(doc);
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function isCJKLocale(lang: string): boolean {
  return (
    lang.startsWith("zh") ||
    lang.startsWith("ja") ||
    lang.startsWith("ko") ||
    lang === "yue"
  );
}

function removeChildrenByLocalName(el: Element, names: string[]): void {
  const toRemove: Element[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i] as Element;
    if (child.localName && names.includes(child.localName)) {
      toRemove.push(child);
    }
  }
  toRemove.forEach((c) => el.removeChild(c));
}

// ─── Main: process PPTX buffer → buffer ─────────────────────────────────────

export async function normalizePptxBuffer(
  input: Buffer | Uint8Array,
  opts: Partial<NormalizeOptions> = {}
): Promise<{ output: Buffer; stats: NormalizeStats }> {
  const options: NormalizeOptions = { ...DEFAULT_OPTIONS, ...opts };
  const zip = await JSZip.loadAsync(input);

  const stats: NormalizeStats = { slides: 0, layouts: 0, masters: 0, themes: 0 };

  if (options.fixTheme) {
    const themeFiles = Object.keys(zip.files).filter((f) =>
      /^ppt\/theme\/theme\d+\.xml$/.test(f)
    );
    for (const themePath of themeFiles) {
      const xml = await zip.file(themePath)!.async("string");
      zip.file(themePath, normalizeThemeFonts(xml, options.targetFont));
      stats.themes++;
    }
  }

  if (options.fixSlides) {
    const slideFiles = Object.keys(zip.files).filter((f) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(f)
    );
    for (const slidePath of slideFiles) {
      const xml = await zip.file(slidePath)!.async("string");
      zip.file(slidePath, normalizeXmlFonts(xml, options));
      stats.slides++;
    }
  }

  if (options.fixLayouts) {
    const layoutFiles = Object.keys(zip.files).filter((f) =>
      /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f)
    );
    for (const layoutPath of layoutFiles) {
      const xml = await zip.file(layoutPath)!.async("string");
      zip.file(layoutPath, normalizeXmlFonts(xml, options));
      stats.layouts++;
    }
  }

  if (options.fixMaster) {
    const masterFiles = Object.keys(zip.files).filter((f) =>
      /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f)
    );
    for (const masterPath of masterFiles) {
      const xml = await zip.file(masterPath)!.async("string");
      zip.file(masterPath, normalizeXmlFonts(xml, options));
      stats.masters++;
    }
  }

  const output = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return { output: Buffer.from(output), stats };
}
