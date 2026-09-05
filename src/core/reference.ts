import { STANDARD_CANON, type CanonBook, type CanonMetadata } from "./canon";

export type NormalizedReference = {
  bookId: string;
  bookName: string;
  chapter: number;
  verseStart: number;
  verseEnd?: number;
};

export type TextRange = { from: number; to: number };
export type ReferenceMatchKind = "exact" | "fuzzy";

type CandidateBase = TextRange & {
  sourceText: string;
  display: string;
  matchKind: ReferenceMatchKind;
};

export type ReferenceCandidate = CandidateBase & {
  kind: "valid";
  reference: NormalizedReference;
};

export type ReferenceIssueCode =
  | "chapter_out_of_range"
  | "verse_out_of_range"
  | "verse_unavailable"
  | "range_reversed"
  | "range_end_out_of_range";

export type InvalidReferenceCandidate = CandidateBase & {
  kind: "invalid";
  bookId: string;
  bookName: string;
  chapter: number;
  verseStart: number;
  verseEnd?: number;
  issue: { code: ReferenceIssueCode; message: string };
};

export type DetectedReference = ReferenceCandidate | InvalidReferenceCandidate;

type CompiledCanon = {
  exactPattern: RegExp;
  aliases: Map<string, CanonBook>;
};

const compiledCanons = new WeakMap<CanonMetadata, CompiledCanon>();
const coordinatePattern = /(\d{1,3}):(\d{1,3})(?:\s*[-–—]\s*(\d{1,3}))?(?=[\s.,;:!?)\]}"'’”…]|[–—](?=\s))/gu;

function normalizeBookName(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/^first(?=\s)/, "1")
    .replace(/^second(?=\s)/, "2")
    .replace(/^third(?=\s)/, "3")
    .replace(/^iii(?=\s)/, "3")
    .replace(/^ii(?=\s)/, "2")
    .replace(/^i(?=\s)/, "1")
    .replace(/[^a-z0-9]/g, "");
}

function escapePattern(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function aliasPattern(alias: string): string {
  return alias.trim().split(/\s+/).map(escapePattern).join("\\s+");
}

function compileCanon(canon: CanonMetadata): CompiledCanon {
  const cached = compiledCanons.get(canon);
  if (cached) return cached;

  const aliases = new Map<string, CanonBook>();
  const spellings: string[] = [];
  for (const book of canon.books) {
    for (const spelling of [book.name, ...book.aliases]) {
      const key = normalizeBookName(spelling);
      const existing = aliases.get(key);
      if (existing && existing.id !== book.id) {
        throw new Error(`Ambiguous canon alias: ${spelling}`);
      }
      aliases.set(key, book);
      spellings.push(spelling);
    }
  }

  const alternatives = [...new Set(spellings)]
    .sort((left, right) => right.length - left.length)
    .map(aliasPattern)
    .join("|");
  const exactPattern = new RegExp(
    `(?<![\\p{L}\\p{N}])(${alternatives})\\.?\\s+(\\d{1,3}):(\\d{1,3})(?:\\s*[-–—]\\s*(\\d{1,3}))?(?=[\\s.,;:!?)\\]}"'’”…]|[–—](?=\\s))`,
    "giu",
  );
  const compiled = { aliases, exactPattern };
  compiledCanons.set(canon, compiled);
  return compiled;
}

function overlaps(range: TextRange, excluded: readonly TextRange[]): boolean {
  return excluded.some((item) => range.from < item.to && range.to > item.from);
}

function likelyUrlContext(text: string, from: number): boolean {
  if (from > 0 && /[/@]/.test(text.charAt(from - 1))) return true;
  const tokenStart = Math.max(text.lastIndexOf(" ", from - 1), text.lastIndexOf("\n", from - 1)) + 1;
  const prefix = text.slice(tokenStart, from);
  return prefix.includes("://")
    || /(?:^|[<(])(?:[a-z][a-z0-9+.-]*:|www\.[^\s]*)$/iu.test(prefix);
}

function displayReference(book: CanonBook, chapter: number, verseStart: number, verseEnd?: number): string {
  return `${book.name} ${chapter}:${verseStart}${verseEnd === undefined ? "" : `-${verseEnd}`}`;
}

function classify(
  book: CanonBook,
  chapter: number,
  verseStart: number,
  verseEnd: number | undefined,
  base: CandidateBase,
): DetectedReference {
  const verseCount = book.verseCounts[chapter - 1];
  const invalid = (code: ReferenceIssueCode, message: string): InvalidReferenceCandidate => ({
    ...base, kind: "invalid", bookId: book.id, bookName: book.name,
    chapter, verseStart, ...(verseEnd === undefined ? {} : { verseEnd }), issue: { code, message },
  });

  if (verseCount === undefined || chapter < 1) {
    return invalid(
      "chapter_out_of_range",
      `${book.name} has chapters 1–${book.verseCounts.length}; chapter ${chapter} does not exist.`,
    );
  }
  if (verseStart < 1 || verseStart > verseCount) {
    return invalid(
      "verse_out_of_range",
      `${book.name} ${chapter} has verses 1–${verseCount}; verse ${verseStart} does not exist.`,
    );
  }
  if (verseEnd !== undefined && verseEnd < verseStart) {
    return invalid("range_reversed", "The ending verse must not come before the starting verse.");
  }
  if (verseEnd !== undefined && verseEnd > verseCount) {
    return invalid(
      "range_end_out_of_range",
      `${book.name} ${chapter} ends at verse ${verseCount}; verse ${verseEnd} does not exist.`,
    );
  }
  const unavailable = book.unavailableVerses?.[chapter] ?? [];
  if (unavailable.some((verse) => verse >= verseStart && verse <= (verseEnd ?? verseStart))) {
    const verse = unavailable.find((value) => value >= verseStart && value <= (verseEnd ?? verseStart));
    return invalid(
      "verse_unavailable",
      `${book.name} ${chapter}:${verse} is not present in this translation's main text.`,
    );
  }

  return {
    ...base,
    kind: "valid",
    reference: {
      bookId: book.id, bookName: book.name, chapter, verseStart,
      ...(verseEnd === undefined ? {} : { verseEnd }),
    },
  };
}

function editDistance(left: string, right: string): number {
  if (left === right) return 0;
  if (!left.length) return right.length;
  if (!right.length) return left.length;
  let previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      current[rightIndex] = Math.min(
        current[rightIndex - 1]! + 1,
        previous[rightIndex]! + 1,
        previous[rightIndex - 1]!
          + (left.charAt(leftIndex - 1) === right.charAt(rightIndex - 1) ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[right.length]!;
}

function ordinal(key: string): string {
  return /^[123]/.test(key) ? key.charAt(0) : "";
}

function fuzzyBookBefore(
  text: string,
  coordinateFrom: number,
  canon: CanonMetadata,
): { book: CanonBook; from: number; source: string } | undefined {
  const windowFrom = Math.max(0, coordinateFrom - 40);
  const fragment = text.slice(windowFrom, coordinateFrom);
  const tokens = [...fragment.matchAll(/(?:[1-3]|[A-Za-z]+)\.?/g)];
  const trailing = tokens.at(-1);
  if (!trailing || !/^[\s.]*$/.test(fragment.slice((trailing.index ?? 0) + trailing[0].length))) return undefined;

  type Match = { book: CanonBook; from: number; source: string; distance: number };
  const matches: Match[] = [];
  for (const token of tokens.slice(-3)) {
    const localFrom = token.index ?? 0;
    const source = fragment.slice(localFrom).trim();
    if (!/^[1-3]|^[A-Z]/.test(source)) continue;
    const key = normalizeBookName(source);
    const letters = key.replace(/^[1-3]/, "");
    if (letters.length < 4) continue;
    const allowedDistance = letters.length > 11 ? 2 : 1;
    for (const book of canon.books) {
      const canonical = normalizeBookName(book.name);
      if (ordinal(key) !== ordinal(canonical)) continue;
      const distance = editDistance(key, canonical);
      if (distance > 0 && distance <= allowedDistance) {
        matches.push({ book, from: windowFrom + localFrom, source, distance });
      }
    }
  }
  matches.sort((left, right) => left.distance - right.distance || right.source.length - left.source.length);
  const best = matches[0];
  if (!best) return undefined;
  if (matches.some((match) => match.book.id !== best.book.id && match.distance === best.distance)) return undefined;
  return best;
}

export function isValidReference(candidate: DetectedReference): candidate is ReferenceCandidate {
  return candidate.kind === "valid";
}

export function scanReferences(
  text: string,
  excluded: readonly TextRange[] = [],
  canon: CanonMetadata = STANDARD_CANON,
): DetectedReference[] {
  const candidates: DetectedReference[] = [];
  const occupied: TextRange[] = [];
  const { aliases, exactPattern } = compileCanon(canon);
  exactPattern.lastIndex = 0;

  for (const match of text.matchAll(exactPattern)) {
    const from = match.index;
    const sourceText = match[0];
    const to = from + sourceText.length;
    occupied.push({ from, to });
    if (overlaps({ from, to }, excluded) || likelyUrlContext(text, from)) continue;
    const book = aliases.get(normalizeBookName(match[1]!));
    if (!book) continue;
    const chapter = Number(match[2]);
    const verseStart = Number(match[3]);
    const verseEnd = match[4] === undefined ? undefined : Number(match[4]);
    candidates.push(classify(book, chapter, verseStart, verseEnd, {
      from, to, sourceText,
      display: displayReference(book, chapter, verseStart, verseEnd),
      matchKind: "exact",
    }));
  }

  coordinatePattern.lastIndex = 0;
  for (const coordinate of text.matchAll(coordinatePattern)) {
    const coordinateFrom = coordinate.index;
    const coordinateTo = coordinateFrom + coordinate[0].length;
    if (overlaps({ from: coordinateFrom, to: coordinateTo }, occupied)) continue;
    const fuzzy = fuzzyBookBefore(text, coordinateFrom, canon);
    if (!fuzzy) continue;
    const from = fuzzy.from;
    const to = coordinateTo;
    if (overlaps({ from, to }, excluded) || likelyUrlContext(text, from)) continue;
    const chapter = Number(coordinate[1]);
    const verseStart = Number(coordinate[2]);
    const verseEnd = coordinate[3] === undefined ? undefined : Number(coordinate[3]);
    candidates.push(classify(fuzzy.book, chapter, verseStart, verseEnd, {
      from, to, sourceText: text.slice(from, to),
      display: displayReference(fuzzy.book, chapter, verseStart, verseEnd),
      matchKind: "fuzzy",
    }));
  }

  return candidates.sort((left, right) => left.from - right.from);
}
