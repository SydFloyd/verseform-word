import { insertedPassageText, type AnnotatedReference } from "../core/freshness";
import { isValidReference, scanReferences, type ReferenceCandidate } from "../core/reference";

export const VFW010_TRANSLATION = {
  id: "VFW-010-FAKE",
  label: "VFW-010 test text",
} as const;

export type ParagraphSnapshot = {
  paragraphId: string;
  text: string;
  revision: number;
};

export type FakePreview = {
  heading: string;
  text: string;
  insertText: string;
};

/**
 * This is deliberately not Scripture. It lets the first Word-host spike prove
 * annotation and replacement behavior without bundling or requesting Bible
 * text before the DBS adapter exists.
 */
export function fakePreviewFor(candidate: ReferenceCandidate): FakePreview {
  const passage = "[Verseform VFW-010 local fake passage. No Scripture text is bundled.]";
  return {
    heading: candidate.display,
    text: `${passage} This host-proof preview is local test data only.`,
    insertText: insertedPassageText(passage, candidate.display, VFW010_TRANSLATION.label),
  };
}

export function candidatesForParagraph(text: string): ReferenceCandidate[] {
  return scanReferences(text)
    .filter(isValidReference)
    .filter((candidate) => !isVfw010GeneratedCitation(text, candidate));
}

function isVfw010GeneratedCitation(text: string, candidate: ReferenceCandidate): boolean {
  return text.charAt(candidate.from - 1) === "("
    && text.startsWith(`, ${VFW010_TRANSLATION.label})`, candidate.to);
}

export function annotationForCandidate(
  annotationId: string,
  paragraph: ParagraphSnapshot,
  candidate: ReferenceCandidate,
): AnnotatedReference {
  return {
    annotationId,
    paragraphId: paragraph.paragraphId,
    paragraphRevision: paragraph.revision,
    paragraphText: paragraph.text,
    sourceText: candidate.sourceText,
    range: { from: candidate.from, to: candidate.to },
    reference: candidate.reference,
    translationId: VFW010_TRANSLATION.id,
  };
}
