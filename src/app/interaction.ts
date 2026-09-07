import { insertedPassageText, type AnnotatedReference } from "../core/freshness";
import { isValidReference, scanReferences, type ReferenceCandidate } from "../core/reference";
import {
  passageDisplay,
  type ScripturePreview,
  type ScriptureProvider,
  type TranslationCatalog,
} from "../core/scripture";

export const VFW010_TRANSLATION = {
  id: "VFW-010-FAKE",
  label: "VFW-010 test text",
} as const;

export type ParagraphSnapshot = {
  paragraphId: string;
  text: string;
  revision: number;
  /** True only when Word confirms a real paragraph follows this one. */
  terminalDelimiter?: boolean;
};

export type SelectionSnapshot = {
  paragraph: ParagraphSnapshot;
  selection: { from: number; to: number };
  /** Present only when Word proves the caret is at the next paragraph start. */
  previousParagraph?: ParagraphSnapshot;
};

export type SelectionReferenceTarget = {
  paragraph: ParagraphSnapshot;
  candidate: ReferenceCandidate;
};

export type FakePreview = ScripturePreview;

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
    attribution: "Local VFW-010 test data. Not Scripture.",
    insertText: insertedPassageText(passage, candidate.display, VFW010_TRANSLATION.label),
    translationId: VFW010_TRANSLATION.id,
    translationName: VFW010_TRANSLATION.label,
    citationLabel: VFW010_TRANSLATION.label,
    cached: false,
  };
}

export class LocalProofScriptureProvider implements ScriptureProvider {
  public async listTranslations(): Promise<TranslationCatalog> {
    return {
      translations: [{
        id: VFW010_TRANSLATION.id,
        citationLabel: VFW010_TRANSLATION.label,
        name: VFW010_TRANSLATION.label,
        attribution: "Local VFW-010 test data. Not Scripture.",
      }],
      cached: false,
    };
  }

  public async getPassage(reference: ReferenceCandidate["reference"]) {
    const candidate: ReferenceCandidate = {
      kind: "valid",
      from: 0,
      to: 0,
      sourceText: passageDisplay(reference),
      display: passageDisplay(reference),
      matchKind: "exact",
      reference,
    };
    const preview = fakePreviewFor(candidate);
    return {
      reference,
      display: passageDisplay(reference),
      translationId: preview.translationId,
      citationLabel: preview.citationLabel,
      translationName: preview.translationName,
      attribution: preview.attribution,
      text: preview.text,
      cached: false,
    };
  }

  public async clearCache(): Promise<void> {}
}

export function candidatesForParagraph(
  text: string,
  options: { terminalDelimiter?: boolean } = {},
): ReferenceCandidate[] {
  // Word's Paragraph.text omits the paragraph mark. Add a detection-only
  // newline when the caller is evaluating a complete Word paragraph so a
  // reference immediately before that real document delimiter is complete.
  const detectionText = options.terminalDelimiter ? `${text}\n` : text;
  return scanReferences(detectionText)
    .filter(isValidReference)
    .filter((candidate) => !isGeneratedCitation(text, candidate));
}

const CURSOR_TRAILING_DELIMITERS = /^[\s.,;:!?)}\]'"”’]*$/u;

/**
 * Resolve one deliberate command target without asking Word to send prose to a
 * provider. A selection must touch exactly one reference. A collapsed caret
 * targets the containing reference or the nearest completed reference followed
 * only by delimiters. At the start of a new paragraph, Word's structurally
 * proven previous paragraph is considered instead.
 */
export function referenceTargetForSelection(
  snapshot: SelectionSnapshot,
): SelectionReferenceTarget | undefined {
  const targetIn = (
    paragraph: ParagraphSnapshot,
    selection: { from: number; to: number },
  ): SelectionReferenceTarget | undefined => {
    // Choosing Fill Scripture is itself a deliberate completion boundary. This
    // differs from passive detection, which must still wait for a typed Word
    // delimiter before it creates an annotation. Word can omit a trailing
    // space from Paragraph.text even though the user just typed one, so the
    // command treats only the paragraph end as complete.
    const candidates = candidatesForParagraph(paragraph.text, {
      terminalDelimiter: true,
    });
    if (selection.from !== selection.to) {
      const intersecting = candidates.filter((candidate) => (
        selection.from < candidate.to && selection.to > candidate.from
      ));
      return intersecting.length === 1
        ? { paragraph, candidate: intersecting[0]! }
        : undefined;
    }

    const caret = selection.from;
    const containing = candidates.filter((candidate) => (
      candidate.from <= caret && caret <= candidate.to
    ));
    if (containing.length === 1) return { paragraph, candidate: containing[0]! };

    const preceding = candidates
      .filter((candidate) => candidate.to <= caret)
      .filter((candidate) => CURSOR_TRAILING_DELIMITERS.test(
        paragraph.text.slice(candidate.to, caret),
      ))
      .sort((left, right) => right.to - left.to);
    return preceding.length ? { paragraph, candidate: preceding[0]! } : undefined;
  };

  const current = targetIn(snapshot.paragraph, snapshot.selection);
  if (current) return current;
  if (snapshot.selection.from !== 0
    || snapshot.selection.to !== 0
    || !snapshot.previousParagraph) return undefined;

  const previous = snapshot.previousParagraph;
  return targetIn(previous, { from: previous.text.length, to: previous.text.length });
}

function isGeneratedCitation(text: string, candidate: ReferenceCandidate): boolean {
  return text.charAt(candidate.from - 1) === "("
    && /^,\s*[A-Za-z0-9_-][A-Za-z0-9 _-]{0,31}\)/u.test(text.slice(candidate.to));
}

export function annotationForCandidate(
  annotationId: string,
  paragraph: ParagraphSnapshot,
  candidate: ReferenceCandidate,
  translationId: string = VFW010_TRANSLATION.id,
): AnnotatedReference {
  return {
    annotationId,
    paragraphId: paragraph.paragraphId,
    paragraphRevision: paragraph.revision,
    paragraphText: paragraph.text,
    sourceText: candidate.sourceText,
    range: { from: candidate.from, to: candidate.to },
    reference: candidate.reference,
    translationId,
  };
}
