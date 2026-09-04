import { describe, expect, it } from "vitest";
import { insertedPassageText, isReferenceFresh, type AnnotatedReference } from "../src/core/freshness";
import { scanReferences } from "../src/core/reference";

describe("local reference detection", () => {
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
    expect(isReferenceFresh(annotation, "Read John 3:16 today.", 4, "ENGWEB")).toBe(false);
  });

  it("creates an editable plain-text passage and citation", () => {
    expect(insertedPassageText("  For God so loved   the world. ", "John 3:16", "NASB"))
      .toBe("For God so loved the world. (John 3:16, NASB)");
    expect(() => insertedPassageText("   ", "John 3:16", "NASB")).toThrow("empty");
  });
});
