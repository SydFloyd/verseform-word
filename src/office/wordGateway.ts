import type { ParagraphSnapshot } from "../app/interaction";
import type { AnnotatedReference } from "../core/freshness";
import type { ReferenceCandidate } from "../core/reference";
import type { AnnotationRemovalResult, ReplacementResult, WordGateway, WordHostHandlers, WordRuntime } from "./gateway";
import { isReplaceableAnnotationState } from "./replacementGuard";

function isMissingOfficeObject(error: unknown): boolean {
  return typeof error === "object"
    && error !== null
    && "code" in error
    && (error as { code?: unknown }).code === "ItemNotFound";
}

/**
 * Office.js adaptation for the WordApi 1.7 walking slice. It purposefully uses
 * no popup-action APIs: those are WordApi 1.8, while VFW-010 supports its
 * explicit, accessible Insert action in the task pane.
 */
type EventRegistration = { remove(): void };

class OfficeWordRuntime implements WordRuntime {
  private stopping: Promise<void> | undefined;

  public constructor(
    // Word.run requires a ClientObject rather than a RequestContext. This
    // document was created in the same original context as every registration.
    private readonly contextAnchor: OfficeExtension.ClientObject,
    private readonly registrations: readonly EventRegistration[],
  ) {}

  public async stop(): Promise<void> {
    if (this.stopping) return this.stopping;
    const task = (async () => {
      let failure: unknown;
      try {
        await Word.run(this.contextAnchor, async (activeContext) => {
          for (const registration of this.registrations) {
            try {
              registration.remove();
            } catch (error) {
              failure ??= error;
            }
          }
          await activeContext.sync();
        });
      } catch (error) {
        failure ??= error;
      }
      if (failure) throw failure;
    })();
    this.stopping = task;
    try {
      await task;
    } finally {
      // Retain the completed promise so later calls are idempotent.
    }
  }
}

export class OfficeWordGateway implements WordGateway {
  public async start(handlers: WordHostHandlers): Promise<WordRuntime> {
    return Word.run(async (context) => {
      const document = context.document;
      const registrations: EventRegistration[] = [];
      try {
        registrations.push(document.onParagraphChanged.add(async (event) => handlers.onParagraphChanged(event.uniqueLocalIds)));
        registrations.push(document.onAnnotationHovered.add(async (event) => handlers.onAnnotationActivated(event.id, "hovered")));
        registrations.push(document.onAnnotationClicked.add(async (event) => handlers.onAnnotationActivated(event.id, "clicked")));
        registrations.push(document.onAnnotationRemoved.add(async (event) => handlers.onAnnotationRemoved(event.ids)));
        await context.sync();
        return new OfficeWordRuntime(document, registrations);
      } catch (error) {
        for (const registration of registrations) {
          try {
            registration.remove();
          } catch {
            // Keep trying all registrations; surface the original start error.
          }
        }
        try {
          await context.sync();
        } catch {
          // Preserve the original registration failure; a later pane can retry.
        }
        throw error;
      }
    });
  }

  public async readParagraph(paragraphId: string): Promise<ParagraphSnapshot | undefined> {
    try {
      return await Word.run(async (context) => {
        const paragraph = context.document.getParagraphByUniqueLocalId(paragraphId);
        paragraph.load("text,uniqueLocalId");
        await context.sync();
        if (paragraph.uniqueLocalId !== paragraphId) return undefined;
        // Word does not expose a monotonically increasing paragraph revision.
        // The kernel stamps events; the complete text is checked again at insert.
        return { paragraphId, text: paragraph.text, revision: 0 };
      });
    } catch (error) {
      if (isMissingOfficeObject(error)) return undefined;
      throw error;
    }
  }

  public async annotate(paragraph: ParagraphSnapshot, candidate: ReferenceCandidate): Promise<string | undefined> {
    return Word.run(async (context) => {
      const wordParagraph = context.document.getParagraphByUniqueLocalId(paragraph.paragraphId);
      wordParagraph.load("text,uniqueLocalId");
      await context.sync();
      if (wordParagraph.uniqueLocalId !== paragraph.paragraphId || wordParagraph.text !== paragraph.text) {
        return undefined;
      }

      // One critique per call makes the returned ID a direct, non-positional map
      // to this candidate. That avoids assuming an ID ordering for duplicates.
      const ids = wordParagraph.insertAnnotations({
        critiques: [{
          colorScheme: "Lavender",
          start: candidate.from,
          length: candidate.to - candidate.from,
        }],
      });
      await context.sync();
      return ids.value.length === 1 ? ids.value[0] : undefined;
    });
  }

  public async removeAnnotations(annotationIds: readonly string[]): Promise<AnnotationRemovalResult> {
    const removedAnnotationIds: string[] = [];
    const missingAnnotationIds: string[] = [];
    const failedAnnotationIds: string[] = [];
    for (const annotationId of annotationIds) {
      try {
        await Word.run(async (context) => {
          context.document.getAnnotationById(annotationId).delete();
          await context.sync();
        });
        removedAnnotationIds.push(annotationId);
      } catch (error) {
        // A user or Word can remove a temporary annotation before our queued
        // cleanup. Continue so one missing ID cannot leak later IDs.
        if (isMissingOfficeObject(error)) {
          missingAnnotationIds.push(annotationId);
        } else {
          failedAnnotationIds.push(annotationId);
        }
      }
    }
    return { removedAnnotationIds, missingAnnotationIds, failedAnnotationIds };
  }

  public async isAnnotationCurrent(annotation: AnnotatedReference): Promise<boolean> {
    try {
      return await Word.run(async (context) => {
        const document = context.document;
        const paragraph = document.getParagraphByUniqueLocalId(annotation.paragraphId);
        const wordAnnotation = document.getAnnotationById(annotation.annotationId);
        const critique = wordAnnotation.critiqueAnnotation;
        const annotationRange = critique.range;
        const rangeParagraphs = annotationRange.paragraphs;

        paragraph.load("text,uniqueLocalId");
        wordAnnotation.load("id,state");
        critique.load("critique");
        annotationRange.load("text");
        rangeParagraphs.load("items/uniqueLocalId");
        await context.sync();

        const critiqueDefinition = critique.critique;
        return wordAnnotation.id === annotation.annotationId
          && paragraph.uniqueLocalId === annotation.paragraphId
          && paragraph.text === annotation.paragraphText
          && annotationRange.text === annotation.sourceText
          && paragraph.text.slice(annotation.range.from, annotation.range.to) === annotation.sourceText
          && isReplaceableAnnotationState(wordAnnotation.state)
          && critiqueDefinition.start === annotation.range.from
          && critiqueDefinition.length === annotation.range.to - annotation.range.from
          && rangeParagraphs.items.length === 1
          && rangeParagraphs.items[0]?.uniqueLocalId === annotation.paragraphId;
      });
    } catch (error) {
      if (isMissingOfficeObject(error)) return false;
      throw error;
    }
  }

  public async replace(annotation: AnnotatedReference, replacementText: string): Promise<ReplacementResult> {
    try {
      return await Word.run(async (context) => {
        const document = context.document;
        const paragraph = document.getParagraphByUniqueLocalId(annotation.paragraphId);
        const wordAnnotation = document.getAnnotationById(annotation.annotationId);
        const critique = wordAnnotation.critiqueAnnotation;
        const annotationRange = critique.range;
        const rangeParagraphs = annotationRange.paragraphs;

        paragraph.load("text,uniqueLocalId");
        wordAnnotation.load("id,state");
        critique.load("critique");
        annotationRange.load("text");
        rangeParagraphs.load("items/uniqueLocalId");
        await context.sync();

        const critiqueDefinition = critique.critique;
        const ownsExactlyOneParagraph = rangeParagraphs.items.length === 1
          && rangeParagraphs.items[0]?.uniqueLocalId === annotation.paragraphId;
        const isExact = wordAnnotation.id === annotation.annotationId
          && paragraph.uniqueLocalId === annotation.paragraphId
          && paragraph.text === annotation.paragraphText
          && annotationRange.text === annotation.sourceText
          && paragraph.text.slice(annotation.range.from, annotation.range.to) === annotation.sourceText
          && isReplaceableAnnotationState(wordAnnotation.state)
          && critiqueDefinition.start === annotation.range.from
          && critiqueDefinition.length === annotation.range.to - annotation.range.from
          && ownsExactlyOneParagraph;
        if (!isExact) return "stale";

        // This is the sole document mutation: Word receives one Replace command
        // on the annotation's own range, rather than a text search for a possibly
        // duplicated reference. Real-host testing must still prove its Undo unit.
        annotationRange.insertText(replacementText, "Replace");
        await context.sync();
        return "replaced";
      });
    } catch (error) {
      return isMissingOfficeObject(error) ? "missing" : "failed";
    }
  }
}
