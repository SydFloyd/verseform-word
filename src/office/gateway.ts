import type { ParagraphSnapshot } from "../app/interaction";
import type { AnnotatedReference } from "../core/freshness";
import type { ReferenceCandidate } from "../core/reference";

export type AnnotationActivation = "hovered" | "clicked";

export type WordHostHandlers = {
  onParagraphChanged(paragraphIds: readonly string[]): Promise<void>;
  onAnnotationActivated(annotationId: string, activation: AnnotationActivation): Promise<void>;
  onAnnotationRemoved(annotationIds: readonly string[]): Promise<void>;
};

export type ReplacementResult = "replaced" | "stale" | "missing" | "ambiguous" | "failed";

export type AnnotationRemovalResult = {
  /** Word confirmed a delete in the current document. */
  removedAnnotationIds: readonly string[];
  /** The exact ID was not present in this current document. */
  missingAnnotationIds: readonly string[];
  failedAnnotationIds: readonly string[];
};

/** A start-specific runtime whose handlers can be removed exactly once. */
export interface WordRuntime {
  stop(): Promise<void>;
}

/**
 * The narrow Word boundary used by the VFW-010 kernel. No implementation of
 * this interface is allowed to call DBS or retain a paragraph after its active
 * operation completes.
 */
export interface WordGateway {
  start(handlers: WordHostHandlers): Promise<WordRuntime>;
  readParagraph(paragraphId: string): Promise<ParagraphSnapshot | undefined>;
  annotate(paragraph: ParagraphSnapshot, candidate: ReferenceCandidate): Promise<string | undefined>;
  removeAnnotations(annotationIds: readonly string[]): Promise<AnnotationRemovalResult>;
  isAnnotationCurrent(annotation: AnnotatedReference): Promise<boolean>;
  replace(annotation: AnnotatedReference, replacementText: string): Promise<ReplacementResult>;
}
