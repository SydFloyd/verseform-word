import type { NormalizedReference, TextRange } from "./reference";

export type AnnotatedReference = {
  annotationId: string;
  paragraphId: string;
  paragraphRevision: number;
  /**
   * The complete local paragraph snapshot that produced this annotation. It is
   * intentionally runtime-only: it is never persisted or sent to a provider.
   */
  paragraphText: string;
  sourceText: string;
  range: TextRange;
  reference: NormalizedReference;
  translationId: string;
};

export function isReferenceFresh(
  annotation: AnnotatedReference,
  currentParagraphText: string,
  currentParagraphRevision: number,
  currentTranslationId: string,
): boolean {
  return annotation.paragraphRevision === currentParagraphRevision
    && annotation.translationId === currentTranslationId
    && annotation.paragraphText === currentParagraphText
    && currentParagraphText.slice(annotation.range.from, annotation.range.to) === annotation.sourceText;
}
export function insertedPassageText(
  passage: string,
  displayReference: string,
  translationLabel: string,
): string {
  const normalizedPassage = passage.replace(/\s+/gu, " ").trim();
  if (!normalizedPassage) throw new Error("The passage is empty.");
  return `${normalizedPassage} (${displayReference}, ${translationLabel})`;
}
