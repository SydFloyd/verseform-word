import type { NormalizedReference } from "./reference";

export const DBS_CATALOG_LIMIT_BYTES = 8 * 1024 * 1024;
export const DBS_CHAPTER_LIMIT_BYTES = 2 * 1024 * 1024;
const MAX_TRANSLATIONS = 6_000;
const MAX_VERSE_ENTRIES = 250;
const translationIdPattern = /^[A-Za-z0-9_-]{1,64}$/u;
const headingMinorWords = new Set([
  "a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "into",
  "of", "on", "or", "the", "through", "to", "versus", "with", "without",
]);

/**
 * DBS BrowserBible section codes, pinned locally from its MIT-licensed
 * BookCodes table. These disambiguate response keys such as S13.1
 * (1 Samuel 3:1), which cannot be split safely without the expected book.
 */
export const DBS_SECTION_BOOK_CODES: Readonly<Record<string, string>> = Object.freeze({
  GEN: "GN", EXO: "EX", LEV: "LV", NUM: "NU", DEU: "DT", JOS: "JS", JDG: "JG", RUT: "RT",
  "1SA": "S1", "2SA": "S2", "1KI": "K1", "2KI": "K2", "1CH": "R1", "2CH": "R2",
  EZR: "ER", NEH: "NH", EST: "ES", JOB: "JB", PSA: "PS", PRO: "PR", ECC: "EC", SNG: "SS",
  ISA: "IS", JER: "JR", LAM: "LM", EZK: "EK", DAN: "DN", HOS: "HO", JOL: "JL", AMO: "AM",
  OBA: "OB", JON: "JH", MIC: "MC", NAM: "NM", HAB: "HK", ZEP: "ZP", HAG: "HG", ZEC: "ZC",
  MAL: "ML", MAT: "MT", MRK: "MK", LUK: "LK", JHN: "JN", ACT: "AC", ROM: "RM", "1CO": "C1",
  "2CO": "C2", GAL: "GL", EPH: "EP", PHP: "PP", COL: "CL", "1TH": "H1", "2TH": "H2",
  "1TI": "T1", "2TI": "T2", TIT: "TT", PHM: "PM", HEB: "HB", JAS: "JM", "1PE": "P1",
  "2PE": "P2", "1JN": "J1", "2JN": "J2", "3JN": "J3", JUD: "JD", REV: "RV",
});

export type Translation = {
  id: string;
  citationLabel: string;
  name: string;
  vernacularName?: string;
  languageCode?: string;
  script?: string;
  year?: string;
  copyright?: string;
  attribution: string;
};

export type TranslationCatalog = {
  translations: Translation[];
  cached: boolean;
};

export type Passage = {
  reference: NormalizedReference;
  display: string;
  translationId: string;
  citationLabel: string;
  translationName: string;
  attribution: string;
  text: string;
  cached: boolean;
};

export type ScripturePreview = {
  heading: string;
  text: string;
  attribution: string;
  insertText: string;
  translationId: string;
  translationName: string;
  citationLabel: string;
  cached: boolean;
};

export interface ScriptureProvider {
  listTranslations(signal?: AbortSignal): Promise<TranslationCatalog>;
  getPassage(
    reference: NormalizedReference,
    translationId: string,
    signal?: AbortSignal,
  ): Promise<Passage>;
  clearCache(): Promise<void>;
}

export interface TranslationPreferenceStore {
  load(): Promise<string | undefined>;
  save(translationId: string): Promise<void>;
}

function record(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function boundedString(value: unknown, maximum: number): string | undefined {
  if (typeof value !== "string" && typeof value !== "number") return undefined;
  const normalized = String(value).replace(/\s+/gu, " ").trim();
  return normalized && normalized.length <= maximum ? normalized : undefined;
}

export function isTranslationId(value: string): boolean {
  return translationIdPattern.test(value);
}

function citationLabel(id: string): string {
  const withoutLanguage = /^[A-Z]{3}[A-Z0-9_-]{2,}$/u.test(id) ? id.slice(3) : id;
  return withoutLanguage.slice(0, 24) || id;
}

function translationAttribution(name: string, id: string, copyright?: string): string {
  return copyright
    ? `${name} (${citationLabel(id)}): ${copyright}`
    : `Scripture quotations from ${name} (${citationLabel(id)}), supplied by Digital Bible Society.`;
}

function encodedLength(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

export function parseDbsCatalog(body: string): Translation[] {
  if (encodedLength(body) > DBS_CATALOG_LIMIT_BYTES) {
    throw new Error("The DBS translation catalog exceeded Verseform's safety limit.");
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("DBS returned a malformed translation catalog.");
  }
  if (!Array.isArray(parsed) || parsed.length > MAX_TRANSLATIONS) {
    throw new Error("DBS returned an invalid translation catalog.");
  }

  const seen = new Set<string>();
  const translations: Translation[] = [];
  for (const value of parsed) {
    const item = record(value);
    const id = boundedString(item?.abbr, 64);
    const name = boundedString(item?.title, 240);
    if (!item || !id || !name || !isTranslationId(id)) {
      throw new Error("DBS returned an invalid translation catalog entry.");
    }
    if (seen.has(id)) continue;
    seen.add(id);
    const copyright = boundedString(item.copyright, 2_000);
    translations.push({
      id,
      citationLabel: citationLabel(id),
      name,
      ...(boundedString(item.title_vernacular, 240) ? { vernacularName: boundedString(item.title_vernacular, 240) } : {}),
      ...(boundedString(item.iso, 16) ? { languageCode: boundedString(item.iso, 16) } : {}),
      ...(boundedString(item.script, 32) ? { script: boundedString(item.script, 32) } : {}),
      ...(boundedString(item.year, 32) ? { year: boundedString(item.year, 32) } : {}),
      ...(copyright ? { copyright } : {}),
      attribution: translationAttribution(name, id, copyright),
    });
  }
  if (!translations.length) throw new Error("The DBS translation catalog was empty.");
  return translations.sort((left, right) => (
    left.name.localeCompare(right.name) || left.id.localeCompare(right.id)
  ));
}

function probableSectionHeading(value: string): boolean {
  if (!value || /[.!?;:]$/u.test(value)) return false;
  const words = value.split(" ");
  if (words.length < 2 || words.length > 12) return false;
  let titleWords = 0;
  for (const word of words) {
    const plain = word.replace(/^[“‘(']+|[”’)',]+$/gu, "");
    if (!plain) return false;
    if (headingMinorWords.has(plain.toLowerCase())) continue;
    if (!/^[\p{Lu}\d][\p{L}\p{M}\d'’()-]*$/u.test(plain)) return false;
    titleWords += 1;
  }
  return titleWords >= 2;
}

export function normalizeDbsVerseText(value: string): string {
  let text = value.replace(/\s+/gu, " ").trim();
  const gluedSentenceBoundaries = [...text.matchAll(/[\p{Ll}\p{N}]\.(?=[\p{Lu}])/gu)];
  const lastBoundary = gluedSentenceBoundaries.at(-1);
  if (lastBoundary?.index !== undefined) {
    const punctuationIndex = lastBoundary.index + lastBoundary[0].length - 1;
    const possibleHeading = text.slice(punctuationIndex + 1);
    if (probableSectionHeading(possibleHeading)) text = text.slice(0, punctuationIndex + 1);
  }
  return text
    .replace(/([,;:!?])(?=[\p{Lu}“‘])/gu, "$1 ")
    .replace(/([\p{Ll}\p{N}”’\])}]\.)(?=[\p{Lu}“‘])/gu, "$1 ");
}

export function parseDbsChapter(
  body: string,
  requestedChapter: number,
  requestedBookId: string,
): Map<number, string> {
  if (encodedLength(body) > DBS_CHAPTER_LIMIT_BYTES) {
    throw new Error("The DBS chapter exceeded Verseform's safety limit.");
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("DBS returned malformed chapter data.");
  }
  if (!Array.isArray(parsed) || parsed.length > MAX_VERSE_ENTRIES) {
    throw new Error("DBS returned invalid chapter data.");
  }
  const sectionBookCode = DBS_SECTION_BOOK_CODES[requestedBookId];
  if (!sectionBookCode) throw new Error("The requested DBS book was not recognized.");
  const expectedPrefix = `${sectionBookCode}${requestedChapter}.`;

  const verses = new Map<number, string[]>();
  let entryCount = 0;
  for (const value of parsed) {
    const item = record(value);
    const entries = item ? Object.entries(item) : [];
    if (!entries.length) {
      if (parsed.length === 1 && item) return new Map();
      throw new Error("DBS returned an invalid verse entry.");
    }
    if (entryCount + entries.length > MAX_VERSE_ENTRIES) {
      throw new Error("DBS returned an invalid verse entry.");
    }
    entryCount += entries.length;
    for (const [key, verseValue] of entries) {
      if (typeof verseValue !== "string" || verseValue.length > 20_000) {
        throw new Error("DBS returned an invalid verse entry.");
      }
      if (!key.startsWith(expectedPrefix)) {
        throw new Error("DBS returned verse data for different coordinates.");
      }
      const match = /^(\d{1,3})([a-z]?)$/u.exec(key.slice(expectedPrefix.length));
      if (!match) throw new Error("DBS returned an invalid verse key.");
      const verse = Number(match[1]);
      if (verse < 1 || verse > 250) throw new Error("DBS returned an invalid verse number.");
      const text = normalizeDbsVerseText(verseValue);
      if (!text) throw new Error("DBS returned an empty verse.");
      const sections = verses.get(verse) ?? [];
      sections.push(text);
      verses.set(verse, sections);
    }
  }
  return new Map([...verses].map(([verse, sections]) => [verse, sections.join(" ")]));
}

function preferredNasb(translations: readonly Translation[]): Translation | undefined {
  return translations.find((translation) => translation.id.toUpperCase() === "ENGNASB")
    ?? translations.find((translation) => translation.citationLabel.toUpperCase() === "NASB")
    ?? translations.find((translation) => /\bNASB\b|New American Standard Bible/iu.test(
      `${translation.name} ${translation.vernacularName ?? ""}`,
    ));
}

export function selectInitialTranslation(
  translations: readonly Translation[],
  savedPreference?: string,
): Translation | undefined {
  return translations.find((translation) => translation.id === savedPreference)
    ?? preferredNasb(translations)
    ?? translations[0];
}

export function passageDisplay(reference: NormalizedReference): string {
  return `${reference.bookName} ${reference.chapter}:${reference.verseStart}${
    reference.verseEnd === undefined ? "" : `-${reference.verseEnd}`
  }`;
}

export function previewForPassage(passage: Passage): ScripturePreview {
  const normalizedPassage = passage.text.replace(/\s+/gu, " ").trim();
  if (!normalizedPassage) throw new Error("The passage is empty.");
  const citation = `(${passage.display}, ${passage.citationLabel})`;
  return {
    heading: `${passage.display} · ${passage.citationLabel}`,
    text: normalizedPassage,
    attribution: passage.attribution,
    insertText: `${normalizedPassage} ${citation}\n${passage.attribution}`,
    translationId: passage.translationId,
    translationName: passage.translationName,
    citationLabel: passage.citationLabel,
    cached: passage.cached,
  };
}
