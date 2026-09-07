import { readFileSync } from "node:fs";
import { performance } from "node:perf_hooks";
import { fileURLToPath } from "node:url";
import { describe, expect, it, vi } from "vitest";
import { candidatesForParagraph } from "../src/app/interaction";
import { STANDARD_CANON } from "../src/core/canon";
import { insertedPassageText, isReferenceFresh, type AnnotatedReference } from "../src/core/freshness";
import { isValidReference, scanReferences, type ReferenceIssueCode } from "../src/core/reference";

type ExpectedReference = {
  sourceText: string;
  display: string;
  bookId: string;
  chapter: number;
  verseStart: number;
  verseEnd?: number;
};

type EnglishCorpus = {
  version: number;
  language: string;
  canonId: string;
  exact: Array<{ input: string; expected: ExpectedReference[] }>;
  fuzzyAccepted: Array<{ input: string; bookId: string }>;
  fuzzyRejected: string[];
  invalid: Array<{ input: string; code: ReferenceIssueCode }>;
  falsePositives: string[];
  delimiters: string[];
  unfinished: string[];
};

const corpusPath = fileURLToPath(new URL("./fixtures/english-reference-corpus-v1.json", import.meta.url));
const corpus = JSON.parse(readFileSync(corpusPath, "utf8")) as EnglishCorpus;

describe("local reference detection", () => {
  it("loads the versioned English-only corpus against the standard canon", () => {
    expect(corpus).toMatchObject({ version: 1, language: "en", canonId: "STANDARD-66" });
    expect(STANDARD_CANON).toMatchObject({ translationId: corpus.canonId, version: 1 });
  });

  it("recognizes every full canonical name and every approved English alias", () => {
    for (const book of STANDARD_CANON.books) {
      for (const spelling of [book.name, ...book.aliases]) {
        expect(scanReferences(`${spelling} 1:1 `), `${book.id}: ${spelling}`).toMatchObject([{
          kind: "valid",
          sourceText: `${spelling} 1:1`,
          reference: { bookId: book.id, chapter: 1, verseStart: 1 },
        }]);
      }
    }
  });

  it("matches the exact, duplicate, ranged, whitespace, and punctuation corpus", () => {
    for (const item of corpus.exact) {
      const actual = scanReferences(item.input);
      expect(actual, item.input).toHaveLength(item.expected.length);
      expect(actual, item.input).toMatchObject(item.expected.map((expected) => ({
        kind: "valid",
        sourceText: expected.sourceText,
        display: expected.display,
        matchKind: "exact",
        reference: {
          bookId: expected.bookId,
          chapter: expected.chapter,
          verseStart: expected.verseStart,
          ...(expected.verseEnd === undefined ? {} : { verseEnd: expected.verseEnd }),
        },
      })));
    }
  });

  it("accepts only conservative, case-signaled fuzzy canonical-book corrections", () => {
    for (const item of corpus.fuzzyAccepted) {
      expect(scanReferences(item.input), item.input).toMatchObject([{
        kind: "valid",
        matchKind: "fuzzy",
        reference: { bookId: item.bookId },
      }]);
    }
    for (const input of corpus.fuzzyRejected) expect(scanReferences(input), input).toEqual([]);
  });

  it("classifies every strict coordinate failure in the corpus", () => {
    for (const item of corpus.invalid) {
      expect(scanReferences(item.input), item.input).toMatchObject([{
        kind: "invalid",
        issue: { code: item.code },
      }]);
    }
  });

  it("requires a completed delimiter and accepts common whitespace and closing punctuation", () => {
    for (const delimiter of corpus.delimiters) {
      expect(scanReferences(`John 3:16${delimiter}`), JSON.stringify(delimiter)).toMatchObject([{
        kind: "valid",
        sourceText: "John 3:16",
      }]);
      expect(scanReferences(`Gensis 1:1${delimiter}`), `fuzzy ${JSON.stringify(delimiter)}`)
        .toMatchObject([{ kind: "valid", matchKind: "fuzzy", sourceText: "Gensis 1:1" }]);
    }
    for (const input of corpus.unfinished) expect(scanReferences(input), input).toEqual([]);
  });

  it("recognizes the two pilot references and treats Word's paragraph mark as a delimiter", () => {
    expect(scanReferences("Hosea 1:1 ")).toMatchObject([{
      kind: "valid",
      sourceText: "Hosea 1:1",
      reference: { bookId: "HOS", chapter: 1, verseStart: 1 },
    }]);
    expect(scanReferences("James 4:17 ")).toMatchObject([{
      kind: "valid",
      sourceText: "James 4:17",
      reference: { bookId: "JAS", chapter: 4, verseStart: 17 },
    }]);
    expect(scanReferences("James 4:17")).toEqual([]);
    expect(candidatesForParagraph("James 4:17", { terminalDelimiter: true })).toMatchObject([{
      sourceText: "James 4:17",
      reference: { bookId: "JAS", chapter: 4, verseStart: 17 },
    }]);
  });

  it("rejects URL, email, compact-identifier, numeric-prose, and excluded-range cases", () => {
    for (const input of corpus.falsePositives) expect(scanReferences(input), input).toEqual([]);
    expect(scanReferences("Keep John 3:16 and Romans 8:28. ", [{ from: 5, to: 14 }]))
      .toMatchObject([{ sourceText: "Romans 8:28" }]);
    expect(candidatesForParagraph("Text (John 3:16, NASB); then John 3:17."))
      .toMatchObject([{ sourceText: "John 3:17" }]);
  });

  it("reports exact UTF-16 offsets inside mixed-direction and surrogate-pair prose", () => {
    const prefix = "שלום 🙏🏽 — ";
    const input = `${prefix}John 3:16 and John 3:17.`;
    const matches = scanReferences(input);
    expect(matches.map(({ from, to, sourceText }) => ({ from, to, sourceText }))).toEqual([
      { from: prefix.length, to: prefix.length + 9, sourceText: "John 3:16" },
      { from: prefix.length + 14, to: prefix.length + 23, sourceText: "John 3:17" },
    ]);
    for (const match of matches) expect(input.slice(match.from, match.to)).toBe(match.sourceText);
  });

  it("scans a 100,000-UTF-16-unit paragraph within the English performance budget", () => {
    const middle = "x".repeat(50_000);
    const input = `Genesis 1:1. ${middle} Romans 8:28. ${middle} Revelation 22:21. `;
    expect(input.length).toBeGreaterThanOrEqual(100_000);
    scanReferences(input);
    const durations = Array.from({ length: 5 }, () => {
      const started = performance.now();
      const matches = scanReferences(input);
      const elapsed = performance.now() - started;
      expect(matches.filter(isValidReference).map((match) => match.reference.bookId))
        .toEqual(["GEN", "ROM", "REV"]);
      return elapsed;
    });
    expect(Math.max(...durations)).toBeLessThan(250);
  });

  it("never calls a provider while evaluating the complete detection corpus", () => {
    const fetchSpy = vi.fn();
    vi.stubGlobal("fetch", fetchSpy);
    try {
      for (const item of corpus.exact) scanReferences(item.input);
      for (const item of corpus.fuzzyAccepted) scanReferences(item.input);
      for (const item of corpus.fuzzyRejected) scanReferences(item);
      for (const item of corpus.invalid) scanReferences(item.input);
      for (const item of corpus.falsePositives) scanReferences(item);
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("waits for a delimiter and accepts exact, ranged, and conservative fuzzy references", () => {
    expect(scanReferences("John 3:16")).toEqual([]);
    expect(scanReferences("John 3:16 ")[0]).toMatchObject({
      kind: "valid",
      display: "John 3:16",
      reference: { bookId: "JHN", chapter: 3, verseStart: 16 },
    });
    expect(scanReferences("Romans 8:28-30.")[0]).toMatchObject({
      kind: "valid",
      reference: { bookId: "ROM", verseStart: 28, verseEnd: 30 },
    });
    expect(scanReferences("Philippianss 4:13\n")[0]).toMatchObject({
      kind: "valid",
      matchKind: "fuzzy",
      reference: { bookId: "PHP" },
    });
  });

  it("classifies strict coordinate failures and rejects representative prose contexts", () => {
    expect(scanReferences("John 3:99 ")[0]).toMatchObject({
      kind: "invalid",
      issue: { code: "verse_out_of_range" },
    });
    expect(scanReferences("Visit https://example.test/John 3:16 today.")).toEqual([]);
    expect(scanReferences("We finished chapter 3:16 yesterday.")).toEqual([]);
  });
});
describe("guarded insertion", () => {
  const annotation: AnnotatedReference = {
    annotationId: "annotation-1",
    paragraphId: "paragraph-1",
    paragraphRevision: 4,
    paragraphText: "Read John 3:16 today.",
    sourceText: "John 3:16",
    range: { from: 5, to: 14 },
    reference: {
      bookId: "JHN",
      bookName: "John",
      chapter: 3,
      verseStart: 16,
    },
    translationId: "ENGNASB",
  };

  it("accepts only the same paragraph slice, revision, and translation", () => {
    expect(isReferenceFresh(annotation, "Read John 3:16 today.", 4, "ENGNASB")).toBe(true);
    expect(isReferenceFresh(annotation, "Read John 3:17 today.", 5, "ENGNASB")).toBe(false);
    expect(isReferenceFresh(annotation, "Please read John 3:16 today.", 4, "ENGNASB")).toBe(false);
    expect(isReferenceFresh(annotation, "Read John 3:16 today.", 4, "ENGWEB")).toBe(false);
  });

  it("creates an editable plain-text passage and citation", () => {
    expect(insertedPassageText("  For God so loved   the world. ", "John 3:16", "NASB"))
      .toBe("For God so loved the world. (John 3:16, NASB)");
    expect(() => insertedPassageText("   ", "John 3:16", "NASB")).toThrow("empty");
  });
});
