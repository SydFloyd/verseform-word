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

export function candidatesForParagraph(text: string): ReferenceCandidate[] {
  return scanReferences(text)
    .filter(isValidReference)
    .filter((candidate) => !isGeneratedCitation(text, candidate));
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
