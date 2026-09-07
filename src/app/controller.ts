import { isReferenceFresh, type AnnotatedReference } from "../core/freshness";
import type { ReferenceCandidate } from "../core/reference";
import {
  passageDisplay,
  previewForPassage,
  selectInitialTranslation,
  type ScripturePreview,
  type ScriptureProvider,
  type Translation,
  type TranslationPreferenceStore,
} from "../core/scripture";
import type { AnnotationRemovalResult, ReplacementResult, WordGateway, WordHostHandlers, WordRuntime } from "../office/gateway";
import { MemoryTranslationPreference } from "../adapters/translationPreference";
import {
  MemoryAnnotationOwnership,
  type AnnotationOwnership,
} from "./annotationOwnership";
import {
  annotationForCandidate,
  candidatesForParagraph,
  fakePreviewFor,
  LocalProofScriptureProvider,
  referenceTargetForSelection,
  type ParagraphSnapshot,
  VFW010_TRANSLATION,
} from "./interaction";

export type VerseformPhase = "starting" | "watching" | "loading-preview" | "preview" | "inserting" | "blocked" | "complete";
export type FocusTarget = "status";
export type CatalogPhase = "loading" | "ready" | "unavailable";

export type VerseformState = {
  phase: VerseformPhase;
  title: string;
  detail: string;
  selectedAnnotationId?: string;
  preview?: ScripturePreview;
  canInsert: boolean;
  canCancel: boolean;
  focusTarget?: FocusTarget;
  catalogPhase?: CatalogPhase;
  translations?: readonly Translation[];
  selectedTranslationId?: string;
  canSelectTranslation?: boolean;
  canClearCache?: boolean;
};

/** Historical aliases retained for the completed VFW-010 proof suite. */
export type Vfw010Phase = VerseformPhase;
export type Vfw010State = VerseformState;

export type StateListener = (state: Readonly<VerseformState>) => void;

export type VerseformControllerOptions = {
  previewOnHover?: boolean;
  /** Reveal the optional pane when an explicit insertion cannot reach Scripture. */
  onInteractionNeedsAttention?: () => void;
};

const initialState: VerseformState = {
  phase: "starting",
  title: "Preparing local detection…",
  detail: "Verseform is registering Word events. No document text leaves Word.",
  canInsert: false,
  canCancel: false,
  catalogPhase: "loading",
  translations: [],
  canSelectTranslation: false,
  canClearCache: false,
};

/**
 * Owns Verseform's finite task-pane lifetime. Paragraph text stays only in this
 * instance's in-memory annotation map and is forgotten on stop. Event handlers
 * capture a runtime generation, so delayed events from a stopped pane do nothing.
 */
export class VerseformController {
  private readonly annotations = new Map<string, AnnotatedReference>();
  /**
   * Bounded, runtime-only evidence for Word Undo. It never enters the browser
   * ledger because that would persist paragraph/source metadata.
   */
  private readonly retiredAnnotations = new Map<string, AnnotatedReference>();
  /**
   * Opaque IDs loaded at this pane start. They carry no prose or range data,
   * but let a post-close resurrection be distinguished from an arbitrary
   * annotation without guessing which occurrence it represents.
   */
  private readonly priorOwnedAnnotationIds = new Set<string>();
  private readonly paragraphRevisions = new Map<string, number>();
  private state: VerseformState = initialState;
  private operation = Promise.resolve();
  private runtime: WordRuntime | undefined;
  private starting: Promise<void> | undefined;
  private stopping: Promise<void> | undefined;
  private active = false;
  private runtimeGeneration = 0;
  private actionGeneration = 0;
  private previewGeneration = 0;
  private previewAbort: AbortController | undefined;
  private inFlightPreview: {
    annotationId: string;
    translationId: string;
    task: Promise<void>;
  } | undefined;
  private catalogAbort: AbortController | undefined;
  private pendingInsert: { token: number; mutationStarted: boolean } | undefined;
  private catalogPhase: CatalogPhase = "loading";
  private translations: Translation[] = [];
  private selectedTranslationId: string | undefined;
  private previewOnHover: boolean;
  private readonly onInteractionNeedsAttention: () => void;

  public constructor(
    private readonly word: WordGateway,
    private readonly onState: StateListener,
    private readonly ownership: AnnotationOwnership = new MemoryAnnotationOwnership(),
    private readonly scripture: ScriptureProvider = new LocalProofScriptureProvider(),
    private readonly preference: TranslationPreferenceStore = new MemoryTranslationPreference(),
    options: VerseformControllerOptions = {},
  ) {
    this.previewOnHover = options.previewOnHover ?? true;
    this.onInteractionNeedsAttention = options.onInteractionNeedsAttention ?? (() => undefined);
  }

  public getState(): Readonly<VerseformState> {
    return this.state;
  }

  /** The optional pane may request hover previews without owning detection. */
  public setPreviewOnHover(enabled: boolean): void {
    this.previewOnHover = enabled;
  }

  /** Idempotently bind one set of handlers for this task-pane lifetime. */
  public async start(): Promise<void> {
    if (this.active && this.runtime) return;
    if (this.starting) return this.starting;

    const generation = ++this.runtimeGeneration;
    this.active = true;
    const task = (async () => {
      try {
        try {
          const owned = await this.ownership.load();
          if (!this.isCurrent(generation)) return;
          this.priorOwnedAnnotationIds.clear();
          for (const entry of owned) this.priorOwnedAnnotationIds.add(entry.annotationId);
          const reconciliation = await this.removeHostAnnotations(
            owned.map((entry) => entry.annotationId),
            "age-missing",
          );
          if (!this.isCurrent(generation)) return;
          if (reconciliation.failedAnnotationIds.length) {
            this.active = false;
            this.reconciliationFailure();
            return;
          }
        } catch {
          if (this.isCurrent(generation)) {
            this.active = false;
            this.runtime = undefined;
            this.ownershipFailure();
          }
          return;
        }

        try {
          const runtime = await this.word.start(this.handlers(generation));
          if (!this.isCurrent(generation)) {
            await this.stopRuntime(runtime);
            return;
          }
          this.runtime = runtime;
          this.publish({
            phase: "watching",
            title: "Watching for completed references",
            detail: "Type a supported English reference followed by a delimiter. Translation access is loading separately.",
            canInsert: false,
            canCancel: false,
          });
          await this.loadCatalog(generation);
        } catch {
          if (this.isCurrent(generation)) {
            this.active = false;
            this.runtime = undefined;
            this.publish({
              phase: "blocked",
              title: "Word events could not start",
              detail: "Verseform did not change your document. Reopen the task pane in a connected Word host with WordApi 1.7.",
              canInsert: false,
              canCancel: false,
              focusTarget: "status",
            });
          }
        }
      } finally {
        this.starting = undefined;
      }
    })();
    this.starting = task;
    return task;
  }

  /**
   * Idempotently remove this lifetime's handlers and temporary annotations.
   * Page teardown invokes this best-effort method; a later pane creates a new
   * controller and runtime generation.
   */
  public async stop(): Promise<void> {
    if (this.stopping) return this.stopping;
    if (!this.active && !this.runtime && !this.starting) return;

    const runtime = this.runtime;
    const startTask = this.starting;
    // Startup owns cross-document reconciliation. This runtime's stop path
    // must not age IDs that were merely missing in another document.
    const annotationIds = [...new Set([
      ...this.annotations.keys(),
      ...this.retiredAnnotations.keys(),
    ])];
    this.runtimeGeneration += 1;
    this.actionGeneration += 1;
    this.previewGeneration += 1;
    this.previewAbort?.abort();
    this.previewAbort = undefined;
    this.inFlightPreview = undefined;
    this.catalogAbort?.abort();
    this.catalogAbort = undefined;
    this.pendingInsert = undefined;
    this.active = false;
    this.runtime = undefined;
    this.annotations.clear();
    this.retiredAnnotations.clear();
    this.priorOwnedAnnotationIds.clear();
    this.paragraphRevisions.clear();

    const task = (async () => {
      await this.removeHostAnnotations(annotationIds, "retain");
      await this.stopRuntime(runtime);
      if (startTask) {
        try {
          await startTask;
        } catch {
          // start() contains its own failure state; teardown must not reject.
        }
      }
    })();
    this.stopping = task;
    try {
      await task;
    } finally {
      if (this.stopping === task) this.stopping = undefined;
    }
  }

  public async clearSelection(): Promise<void> {
    if (this.pendingInsert?.mutationStarted) return;
    if (this.pendingInsert) {
      this.actionGeneration += 1;
      this.pendingInsert = undefined;
    }
    this.cancelPreviewRequest();
    if (!this.state.selectedAnnotationId) return;
    this.publish({
      phase: "watching",
      title: "Reference selection cleared",
      detail: "No text changed. Activate a current Verseform annotation whenever you are ready.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  public async selectTranslation(translationId: string): Promise<void> {
    if (this.pendingInsert) return;
    const selected = this.translations.find((translation) => translation.id === translationId);
    if (!selected || selected.id === this.selectedTranslationId) return;
    this.selectedTranslationId = selected.id;
    this.cancelPreviewRequest();
    this.publish({
      phase: "watching",
      title: `${selected.citationLabel} selected`,
      detail: `${selected.name} will be used for the next preview. No document text was sent.`,
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
    try {
      await this.preference.save(selected.id);
    } catch {
      if (this.selectedTranslationId === selected.id) {
        this.publish({
          phase: "watching",
          title: `${selected.citationLabel} selected for this pane`,
          detail: "The translation works now, but Verseform could not remember it on this device.",
          canInsert: false,
          canCancel: false,
          focusTarget: "status",
        });
      }
    }
  }

  public async clearScriptureCache(): Promise<void> {
    if (this.pendingInsert) return;
    this.cancelPreviewRequest();
    try {
      await this.scripture.clearCache();
      this.publish({
        phase: "watching",
        title: "Local Scripture cache cleared",
        detail: "Saved provider responses were removed. Your translation preference was kept.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
    } catch {
      this.publish({
        phase: "watching",
        title: "Local Scripture cache was not cleared",
        detail: "No document text changed. Reopen the task pane and try again.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
    }
  }

  public async insertSelected(): Promise<void> {
    if (this.pendingInsert) return;
    const selectedId = this.state.selectedAnnotationId;
    const annotation = selectedId ? this.annotations.get(selectedId) : undefined;
    const preview = this.state.preview;
    if (!this.active || !annotation || !preview) {
      this.publish({
        phase: "watching",
        title: "Choose a reference first",
        detail: "Activate a Verseform annotation in Word before inserting.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      return;
    }
    if (!this.selectedTranslationId
      || preview.translationId !== this.selectedTranslationId
      || annotation.translationId !== preview.translationId) {
      this.publish({
        phase: "watching",
        title: "Translation changed before insertion",
        detail: "No text changed. Activate the reference again to load the selected translation.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      return;
    }

    const generation = this.runtimeGeneration;
    const actionToken = ++this.actionGeneration;
    this.pendingInsert = { token: actionToken, mutationStarted: false };
    this.publish({
      phase: "inserting",
      title: "Checking the reference one last time…",
      detail: "Verseform will replace only this exact, unchanged annotation.",
      selectedAnnotationId: annotation.annotationId,
      preview,
      canInsert: false,
      canCancel: true,
    });
    try {
      await this.serial(async () => {
        if (!this.isCurrent(generation) || !this.isActionCurrent(actionToken)) return;
        const current = await this.word.readParagraph(annotation.paragraphId);
        if (!this.isCurrent(generation) || !this.isActionCurrent(actionToken)) return;
        if (!current || !this.isFresh(annotation, this.withKnownRevision(current), preview.translationId)) {
          await this.rejectStale(annotation.annotationId, actionToken);
          return;
        }

        const pendingInsert = this.pendingInsert;
        if (!pendingInsert || !this.isActionCurrent(actionToken)) return;
        pendingInsert.mutationStarted = true;
        this.publish({ ...this.state, canCancel: false });

        let result: ReplacementResult;
        try {
          result = await this.word.replace(annotation, preview.insertText);
        } catch {
          const cleaned = await this.discard(annotation.annotationId);
          if (this.isCurrent(generation) && this.isActionCurrent(actionToken)) {
            this.finishAction(actionToken);
            if (cleaned) this.hostFailure("Word did not confirm the replacement");
            else this.cleanupFailure("Verseform could not clear the failed annotation");
          }
          return;
        }
        if (!this.isCurrent(generation) || !this.isActionCurrent(actionToken)) return;
        await this.handleReplacement(result, annotation, actionToken);
      });
    } catch {
      if (this.isCurrent(generation) && this.isActionCurrent(actionToken)) {
        this.finishAction(actionToken);
        this.hostFailure("Word could not complete the insertion");
      }
    }
  }

  /**
   * Pane-independent command path for hosts that keep native annotation UI
   * dormant while the task pane is hidden. The command captures one local
   * selection target, refreshes that paragraph, then reuses the exact
   * annotation/freshness/DBS/replace path used by annotation activation.
   */
  public async fillAtSelection(): Promise<void> {
    if (this.pendingInsert) return;
    const generation = this.runtimeGeneration;
    let targetAnnotationId: string | undefined;
    let previewTask: Promise<void> | undefined;

    try {
      await this.serial(async () => {
        if (!this.isCurrent(generation)) return;
        const selection = await this.word.readSelection();
        if (!this.isCurrent(generation)) return;
        const target = selection ? referenceTargetForSelection(selection) : undefined;
        if (!target) {
          this.publish({
            phase: "watching",
            title: "Place the cursor by one reference",
            detail: "No text changed. Put the cursor in or immediately after one completed reference, then choose Fill Scripture again.",
            canInsert: false,
            canCancel: false,
            focusTarget: "status",
          });
          this.revealInteractionFailure();
          return;
        }

        // Do not depend on a background paragraph event having been delivered
        // or on Word having painted its temporary annotation while hidden.
        await this.refreshParagraph(target.paragraph.paragraphId, generation, true);
        if (!this.isCurrent(generation)) return;
        const annotation = [...this.annotations.values()].find((item) => (
          item.paragraphId === target.paragraph.paragraphId
          && item.range.from === target.candidate.from
          && item.range.to === target.candidate.to
          && item.sourceText === target.candidate.sourceText
        ));
        if (!annotation) {
          this.publish({
            phase: "watching",
            title: "Reference changed before filling",
            detail: "No text changed. Place the cursor by the completed reference and try again.",
            canInsert: false,
            canCancel: false,
            focusTarget: "status",
          });
          this.revealInteractionFailure();
          return;
        }

        targetAnnotationId = annotation.annotationId;
        previewTask = (await this.activate(annotation.annotationId, generation, true))?.previewTask;
      });
      await previewTask;
      if (targetAnnotationId && this.isInsertReadyFor(targetAnnotationId)) {
        await this.insertSelected();
      }
    } catch {
      if (this.isCurrent(generation)) {
        this.hostFailure("Word could not read the reference at the cursor");
        this.revealInteractionFailure();
      }
    }
  }

  private handlers(generation: number): WordHostHandlers {
    return {
      onParagraphChanged: async (paragraphIds) => {
        await this.runEvent(generation, async () => {
          for (const paragraphId of paragraphIds) {
            if (!this.isCurrent(generation)) return;
            await this.refreshParagraph(paragraphId, generation);
          }
        });
      },
      onParagraphBoundary: async (paragraphIds) => {
        await this.runEvent(generation, async () => {
          for (const paragraphId of paragraphIds) {
            if (!this.isCurrent(generation)) return;
            await this.refreshParagraph(paragraphId, generation);
          }
        });
      },
      onAnnotationActivated: async (annotationId, activation) => {
        if (activation === "hovered" && !this.previewOnHover) return;
        let previewTask: Promise<void> | undefined;
        await this.runEvent(generation, async () => {
          if (this.pendingInsert) return;
          if (activation === "clicked" && this.isInsertReadyFor(annotationId)) return;
          previewTask = (await this.activate(
            annotationId,
            generation,
            activation === "clicked",
          ))?.previewTask;
        });
        await previewTask;
        if (activation === "clicked" && this.isInsertReadyFor(annotationId)) {
          await this.insertSelected();
        }
      },
      onAnnotationRemoved: async (annotationIds) => {
        await this.runEvent(generation, async () => {
          for (const annotationId of annotationIds) {
            const annotation = this.annotations.get(annotationId);
            this.annotations.delete(annotationId);
            // An annotation-removed event can precede its paragraph-changed
            // event. Retain its exact runtime record so an Undo resurrection
            // is never treated as an arbitrary/unknown marker. WordApi 1.7
            // exposes no proof that a removed annotation can never return.
            if (annotation) await this.retireAnnotation(annotation);
            if (this.state.selectedAnnotationId === annotationId) {
              this.publish({
                phase: "watching",
                title: "Reference annotation was removed",
                detail: "No text changed. Edit or complete the reference again to make it available.",
                canInsert: false,
                canCancel: false,
                focusTarget: "status",
              });
            }
          }
        });
      },
    };
  }

  private async runEvent(generation: number, task: () => Promise<void>): Promise<void> {
    if (!this.isCurrent(generation)) return;
    try {
      await this.serial(async () => {
        if (this.isCurrent(generation)) await task();
      });
    } catch {
      if (this.isCurrent(generation)) this.hostFailure("Word could not refresh this reference");
    }
  }

  private async refreshParagraph(
    paragraphId: string,
    generation: number,
    commandCompletesParagraphEnd = false,
  ): Promise<void> {
    const paragraph = await this.word.readParagraph(paragraphId);
    if (!paragraph || !this.isCurrent(generation)) return;

    const revision = (this.paragraphRevisions.get(paragraphId) ?? 0) + 1;
    this.paragraphRevisions.set(paragraphId, revision);
    const snapshot: ParagraphSnapshot = {
      ...paragraph,
      revision,
      terminalDelimiter: paragraph.terminalDelimiter || commandCompletesParagraphEnd,
    };

    const oldAnnotations = [...this.annotations.values()]
      .filter((annotation) => annotation.paragraphId === paragraphId);
    const oldIds = oldAnnotations.map((annotation) => annotation.annotationId);
    // Word Undo can resurrect annotations that Word removed while refreshing a
    // changed paragraph. Retire every exact record before requesting cleanup:
    // its opaque ID remains owned and its runtime-only snapshot can be checked
    // before it is ever actionable again. Do not release these IDs merely
    // because the current cleanup reports success or missing.
    for (const annotation of oldAnnotations) {
      this.annotations.delete(annotation.annotationId);
      await this.retireAnnotation(annotation);
    }
    const cleanup = await this.removeHostAnnotations(oldIds, "retain");
    const failedIds = new Set(cleanup.failedAnnotationIds);
    for (const annotation of oldAnnotations) {
      if (!failedIds.has(annotation.annotationId)) continue;
      this.retiredAnnotations.delete(annotation.annotationId);
      this.annotations.set(annotation.annotationId, annotation);
    }
    if (cleanup.failedAnnotationIds.length) {
      this.cleanupFailure("Verseform could not clear previous annotations");
      return;
    }
    if (this.state.selectedAnnotationId && oldIds.includes(this.state.selectedAnnotationId)) {
      this.cancelPreviewRequest();
      this.publish({
        phase: "watching",
        title: "Reference changed",
        detail: "Verseform cleared the previous action. It will mark a completed reference again when it is fresh.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
    }
    if (!this.isCurrent(generation)) return;

    const candidates = candidatesForParagraph(snapshot.text, {
      terminalDelimiter: snapshot.terminalDelimiter,
    });
    let marked = 0;
    for (const candidate of candidates) {
      if (await this.restoreRetiredAnnotation(snapshot, candidate)) {
        if (!this.isCurrent(generation)) return;
        marked += 1;
        continue;
      }
      let reservation: string | undefined;
      try {
        reservation = await this.ownership.reserve();
      } catch {
        if (this.isCurrent(generation)) this.ownershipFailure();
        return;
      }
      if (!reservation) {
        if (this.isCurrent(generation)) {
          // A retired ID can reappear when Word Undo restores its annotation.
          // Releasing it merely to create another marker would make that
          // resurrected marker unowned, so capacity exhaustion is fail-closed.
          this.cleanupFailure("Verseform could not reserve temporary annotation ownership");
        }
        return;
      }
      if (!this.isCurrent(generation)) {
        await this.cancelReservation(reservation);
        return;
      }

      let annotationId: string | undefined;
      try {
        annotationId = await this.word.annotate(snapshot, candidate);
      } catch {
        await this.cancelReservation(reservation);
        throw new Error("Word could not create this annotation");
      }
      if (!annotationId) {
        await this.cancelReservation(reservation);
        continue;
      }

      let recorded = false;
      try {
        recorded = await this.ownership.commit(reservation, annotationId);
      } catch {
        recorded = false;
      }
      if (!recorded) {
        await this.cancelReservation(reservation);
        await this.removeHostAnnotations([annotationId], "retain");
        if (this.isCurrent(generation)) {
          this.cleanupFailure("Verseform could not record temporary annotation ownership");
        }
        return;
      }
      if (!this.isCurrent(generation)) {
        // This ID committed after a stop raced the create path. The stopped
        // runtime has no safe paragraph metadata to retain, but the opaque
        // ledger must still own it if Word later resurrects the marker.
        await this.removeHostAnnotations([annotationId], "retain");
        return;
      }
      this.annotations.set(annotationId, annotationForCandidate(
        annotationId,
        snapshot,
        candidate,
        this.selectedTranslationId ?? "DBS-PENDING",
      ));
      marked += 1;
    }

    if (!this.isCurrent(generation)) return;
    if (marked) {
      this.publish({
        phase: "watching",
        title: marked === 1 ? "Reference ready" : `${marked} references ready`,
        detail: this.selectedTranslationId
          ? "Activate a Verseform annotation in Word to request its passage from DBS."
          : "The reference was detected locally. Scripture text needs a connection before preview or insertion.",
        canInsert: false,
        canCancel: false,
      });
    } else if (oldIds.length) {
      this.publish({
        phase: "watching",
        title: "Watching for completed references",
        detail: "The previous annotation is no longer current. Type a supported reference followed by a delimiter.",
        canInsert: false,
        canCancel: false,
      });
    }
  }

  private async restoreRetiredAnnotation(
    snapshot: ParagraphSnapshot,
    candidate: ReferenceCandidate,
  ): Promise<boolean> {
    const retired = [...this.retiredAnnotations.values()].filter((annotation) => (
      annotation.paragraphId === snapshot.paragraphId
      && annotation.paragraphText === snapshot.text
      && annotation.range.from === candidate.from
      && annotation.range.to === candidate.to
      && annotation.sourceText === candidate.sourceText
    ));
    for (const annotation of retired) {
      if (!await this.word.isAnnotationCurrent(annotation)) continue;
      this.retiredAnnotations.delete(annotation.annotationId);
      this.annotations.set(annotation.annotationId, {
        ...annotation,
        paragraphText: snapshot.text,
        paragraphRevision: snapshot.revision,
      });
      return true;
    }
    return false;
  }

  /**
   * Word can deliver an activation after a paragraph refresh has already
   * considered a different resurrected marker. Only an ID retained by this
   * runtime may take this path, and the gateway rechecks the full native
   * identity before it becomes actionable.
   */
  private async restoreRetiredForActivation(
    annotationId: string,
    generation: number,
  ): Promise<AnnotatedReference | undefined> {
    const retired = this.retiredAnnotations.get(annotationId);
    if (!retired || !this.isCurrent(generation)) return undefined;
    const current = await this.word.readParagraph(retired.paragraphId);
    if (!current || !this.isCurrent(generation)) return undefined;
    const recovered: AnnotatedReference = {
      ...retired,
      paragraphText: current.text,
      paragraphRevision: this.withKnownRevision(current).revision,
    };
    if (!await this.word.isAnnotationCurrent(recovered) || !this.isCurrent(generation)) return undefined;

    // If refresh already made a newer exact marker before Word delivered this
    // resurrected ID, remove only that known duplicate. Both IDs stay retired
    // in our bounded ownership set, because deleting the newer marker is also
    // potentially undoable. A failed delete is not silently papered over.
    const newerDuplicates = [...this.annotations.values()].filter((annotation) => (
      annotation.annotationId !== annotationId
      && annotation.paragraphId === recovered.paragraphId
      && annotation.range.from === recovered.range.from
      && annotation.range.to === recovered.range.to
      && annotation.sourceText === recovered.sourceText
    ));
    for (const duplicate of newerDuplicates) {
      this.annotations.delete(duplicate.annotationId);
      await this.retireAnnotation(duplicate);
    }
    const cleanup = await this.removeHostAnnotations(
      newerDuplicates.map((annotation) => annotation.annotationId),
      "retain",
    );
    if (cleanup.failedAnnotationIds.length) {
      for (const duplicate of newerDuplicates) {
        if (!cleanup.failedAnnotationIds.includes(duplicate.annotationId)) continue;
        this.retiredAnnotations.delete(duplicate.annotationId);
        this.annotations.set(duplicate.annotationId, duplicate);
      }
      this.cleanupFailure("Verseform could not clear a duplicate temporary annotation");
      return undefined;
    }

    this.retiredAnnotations.delete(annotationId);
    this.annotations.set(annotationId, recovered);
    return recovered;
  }

  private async activate(
    annotationId: string,
    generation: number,
    revealOnFailure: boolean,
  ): Promise<{ previewTask: Promise<void> } | undefined> {
    let annotation = this.annotations.get(annotationId);
    if (!annotation) annotation = await this.restoreRetiredForActivation(annotationId, generation);
    if (!annotation) {
      if (this.state.phase === "blocked") return;
      const wasOwnedBeforeThisPane = this.priorOwnedAnnotationIds.has(annotationId);
      this.publish({
        phase: "watching",
        title: wasOwnedBeforeThisPane
          ? "A previous temporary annotation returned"
          : "That annotation is not available",
        detail: wasOwnedBeforeThisPane
          ? "Verseform recognizes this earlier annotation ID but cannot safely map it after reopening. No text changed. Complete the reference again to create a current annotation."
          : "Verseform will not guess which duplicate reference you meant. Complete or activate a current annotation instead.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      return undefined;
    }

    const translation = this.translations.find((item) => item.id === this.selectedTranslationId);
    if (!translation) {
      this.publish({
        phase: "watching",
        title: "Scripture text needs a connection",
        detail: "Reference detection remains local. Reopen the task pane when DBS is available to preview or insert Scripture.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      if (revealOnFailure) this.revealInteractionFailure();
      return undefined;
    }

    const inFlight = this.inFlightPreview;
    if (inFlight
      && inFlight.annotationId === annotationId
      && inFlight.translationId === translation.id
      && this.state.selectedAnnotationId === annotationId) {
      return { previewTask: inFlight.task };
    }

    const current = await this.word.readParagraph(annotation.paragraphId);
    if (!this.isCurrent(generation)) return;
    const translatedAnnotation = { ...annotation, translationId: translation.id };
    if (!current || !this.isFresh(
      translatedAnnotation,
      this.withKnownRevision(current),
      translation.id,
    )) {
      await this.rejectStale(annotation.annotationId);
      return undefined;
    }
    this.annotations.set(annotationId, translatedAnnotation);
    this.cancelPreviewRequest();
    const requestToken = ++this.previewGeneration;
    const abort = new AbortController();
    this.previewAbort = abort;
    this.publish({
      phase: "loading-preview",
      title: `Loading ${translation.citationLabel}…`,
      detail: `Requesting only ${passageDisplay(annotation.reference)} coordinates from DBS. Document prose stays in Word.`,
      selectedAnnotationId: annotationId,
      canInsert: false,
      canCancel: true,
    });
    const previewTask = this.loadPreview(
      translatedAnnotation,
      translation,
      requestToken,
      generation,
      abort.signal,
      revealOnFailure,
    );
    this.inFlightPreview = {
      annotationId,
      translationId: translation.id,
      task: previewTask,
    };
    void previewTask.finally(() => {
      if (this.inFlightPreview?.task === previewTask) this.inFlightPreview = undefined;
    });
    return { previewTask };
  }

  private isInsertReadyFor(annotationId: string): boolean {
    return this.active
      && this.state.phase === "preview"
      && this.state.selectedAnnotationId === annotationId
      && this.state.preview !== undefined
      && this.state.canInsert;
  }

  private isFresh(
    annotation: AnnotatedReference,
    current: ParagraphSnapshot,
    expectedTranslationId = this.selectedTranslationId ?? "",
  ): boolean {
    return isReferenceFresh(annotation, current.text, current.revision, expectedTranslationId);
  }

  private async handleReplacement(
    result: ReplacementResult,
    annotation: AnnotatedReference,
    actionToken: number,
  ): Promise<void> {
    if (result === "replaced") {
      const insertedPreview = this.state.preview;
      const localProof = insertedPreview?.translationId === VFW010_TRANSLATION.id;
      this.annotations.delete(annotation.annotationId);
      await this.retireAnnotation(annotation);
      this.finishAction(actionToken);
      this.publish({
        phase: "complete",
        title: localProof
          ? "VFW-010 test replacement inserted"
          : `${insertedPreview?.citationLabel ?? "Scripture"} passage inserted`,
        detail: localProof
          ? "Use Word’s Undo command to restore the reference. The inserted text is local test data, not Scripture."
          : "The passage and editable translation citation were inserted together. Use Word’s Undo command to restore the reference.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      return;
    }

    const cleaned = await this.discard(annotation.annotationId);
    if (!this.isActionCurrent(actionToken)) return;
    this.finishAction(actionToken);
    if (!cleaned) {
      this.cleanupFailure("Verseform could not clear the rejected annotation");
      return;
    }
    if (result === "ambiguous") {
      this.publish({
        phase: "watching",
        title: "Verseform refused an ambiguous replacement",
        detail: "No text changed. Activate one current annotation and try again.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      return;
    }
    if (result === "failed") {
      this.hostFailure("Word did not confirm the replacement");
      return;
    }
    this.publish({
      phase: "watching",
      title: result === "missing" ? "Reference annotation is gone" : "Writing changed before insertion",
      detail: "No text changed. Complete or activate the current reference again before inserting.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  private async rejectStale(annotationId: string, actionToken?: number): Promise<void> {
    this.cancelPreviewRequest();
    const cleaned = await this.discard(annotationId);
    if (actionToken !== undefined) {
      if (!this.isActionCurrent(actionToken)) return;
      this.finishAction(actionToken);
    }
    if (!cleaned) {
      this.cleanupFailure("Verseform could not clear the stale annotation");
      return;
    }
    this.publish({
      phase: "watching",
      title: "Writing changed before insertion",
      detail: "No text changed. Complete or activate the current reference again before inserting.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  private async retireAnnotation(annotation: AnnotatedReference): Promise<void> {
    // Keep the exact runtime snapshot for an Undo-resurrected ID. Do not evict
    // an ID while Word may still restore it: WordApi 1.7 gives us no owner tag
    // with which to recover an evicted annotation safely.
    this.retiredAnnotations.set(annotation.annotationId, annotation);
  }

  private async discard(annotationId: string): Promise<boolean> {
    const annotation = this.annotations.get(annotationId);
    if (annotation) {
      this.annotations.delete(annotationId);
      await this.retireAnnotation(annotation);
    }
    const cleanup = await this.removeHostAnnotations([annotationId], "retain");
    if (cleanup.failedAnnotationIds.includes(annotationId) && annotation) {
      // A failed removal remains an active cleanup obligation. Restoring it to
      // the active map makes the next paragraph refresh retry before it can
      // create another marker over the still-visible native annotation.
      this.retiredAnnotations.delete(annotationId);
      this.annotations.set(annotationId, annotation);
    }
    return cleanup.failedAnnotationIds.length === 0;
  }

  private async removeHostAnnotations(
    annotationIds: readonly string[],
    ownershipDisposition: "retain" | "age-missing" = "retain",
  ): Promise<AnnotationRemovalResult> {
    if (!annotationIds.length) {
      return { removedAnnotationIds: [], missingAnnotationIds: [], failedAnnotationIds: [] };
    }
    try {
      const result = await this.word.removeAnnotations(annotationIds);
      if (ownershipDisposition === "age-missing") {
        await this.ownership.recordRemoval({
          removedAnnotationIds: [],
          missingAnnotationIds: result.missingAnnotationIds,
        });
      }
      return result;
    } catch {
      return { removedAnnotationIds: [], missingAnnotationIds: [], failedAnnotationIds: [...annotationIds] };
    }
  }

  private async cancelReservation(reservationId: string): Promise<void> {
    try {
      await this.ownership.cancel(reservationId);
    } catch {
      // The bounded reservation expires without a Word annotation. The caller
      // still reports the original gateway failure through its normal path.
    }
  }

  private async stopRuntime(runtime: WordRuntime | undefined): Promise<void> {
    if (!runtime) return;
    try {
      await runtime.stop();
    } catch {
      // Teardown must not create an unhandled promise rejection.
    }
  }

  private withKnownRevision(paragraph: ParagraphSnapshot): ParagraphSnapshot {
    return { ...paragraph, revision: this.paragraphRevisions.get(paragraph.paragraphId) ?? 0 };
  }

  private async loadCatalog(generation: number): Promise<void> {
    const abort = new AbortController();
    this.catalogAbort = abort;
    try {
      const [catalog, savedPreference] = await Promise.all([
        this.scripture.listTranslations(abort.signal),
        this.preference.load().catch(() => undefined),
      ]);
      if (!this.isCurrent(generation) || abort.signal.aborted) return;
      this.translations = [...catalog.translations];
      const selected = selectInitialTranslation(this.translations, savedPreference);
      this.selectedTranslationId = selected?.id;
      this.catalogPhase = selected ? "ready" : "unavailable";
      this.publish({
        phase: "watching",
        title: selected ? `${selected.citationLabel} is ready` : "No authorized translations are available",
        detail: selected
          ? `${selected.name} is selected${catalog.cached ? " from the bounded local cache" : " from DBS"}. Detection remains local.`
          : "Reference detection remains local, but preview and insertion need an authorized DBS translation.",
        canInsert: false,
        canCancel: false,
      });
    } catch (error) {
      if (!this.isCurrent(generation) || abort.signal.aborted) return;
      this.translations = [];
      this.selectedTranslationId = undefined;
      this.catalogPhase = "unavailable";
      this.publish({
        phase: "watching",
        title: "Scripture service is unavailable",
        detail: error instanceof Error
          ? `${error.message} Reference detection remains local; no document prose was sent.`
          : "Reference detection remains local. Reopen the task pane when DBS is available.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
    } finally {
      if (this.catalogAbort === abort) this.catalogAbort = undefined;
    }
  }

  private async loadPreview(
    annotation: AnnotatedReference,
    translation: Translation,
    requestToken: number,
    generation: number,
    signal: AbortSignal,
    revealOnFailure: boolean,
  ): Promise<void> {
    try {
      const passage = await this.scripture.getPassage(
        annotation.reference,
        translation.id,
        signal,
      );
      if (!this.isPreviewCurrent(requestToken, generation, annotation.annotationId, translation.id)) return;
      if (passage.translationId !== translation.id
        || passage.display !== passageDisplay(annotation.reference)) {
        throw new Error("DBS returned a passage for different coordinates.");
      }
      await this.serial(async () => {
        if (!this.isPreviewCurrent(requestToken, generation, annotation.annotationId, translation.id)) return;
        const current = await this.word.readParagraph(annotation.paragraphId);
        if (!this.isPreviewCurrent(requestToken, generation, annotation.annotationId, translation.id)) return;
        if (!current || !this.isFresh(annotation, this.withKnownRevision(current), translation.id)) {
          await this.rejectStale(annotation.annotationId);
          return;
        }
        const candidate: ReferenceCandidate = {
          kind: "valid",
          from: annotation.range.from,
          to: annotation.range.to,
          sourceText: annotation.sourceText,
          display: passageDisplay(annotation.reference),
          matchKind: "exact",
          reference: annotation.reference,
        };
        const preview = translation.id === VFW010_TRANSLATION.id
          ? fakePreviewFor(candidate)
          : previewForPassage(passage);
        this.previewAbort = undefined;
        this.publish({
          phase: "preview",
          title: "Preview ready",
          detail: `${translation.name}${passage.cached ? " · local cache" : " · DBS"} · click the marked reference to insert`,
          selectedAnnotationId: annotation.annotationId,
          preview,
          canInsert: true,
          canCancel: true,
        });
      });
    } catch (error) {
      if (signal.aborted || !this.isPreviewCurrent(
        requestToken,
        generation,
        annotation.annotationId,
        translation.id,
      )) return;
      this.previewAbort = undefined;
      this.publish({
        phase: "watching",
        title: "Scripture text is unavailable",
        detail: error instanceof Error
          ? `${error.message} No document text changed.`
          : "DBS could not provide this passage. No document text changed.",
        canInsert: false,
        canCancel: false,
        focusTarget: "status",
      });
      if (revealOnFailure) this.revealInteractionFailure();
    }
  }

  private revealInteractionFailure(): void {
    try {
      this.onInteractionNeedsAttention();
    } catch {
      // A host that cannot show the optional pane must not turn a provider
      // failure into an unhandled event or change document text.
    }
  }

  private isPreviewCurrent(
    requestToken: number,
    generation: number,
    annotationId: string,
    translationId: string,
  ): boolean {
    return this.isCurrent(generation)
      && requestToken === this.previewGeneration
      && this.state.selectedAnnotationId === annotationId
      && this.selectedTranslationId === translationId;
  }

  private cancelPreviewRequest(): void {
    this.previewGeneration += 1;
    this.previewAbort?.abort();
    this.previewAbort = undefined;
    this.inFlightPreview = undefined;
  }

  private isCurrent(generation: number): boolean {
    return this.active && generation === this.runtimeGeneration;
  }

  private isActionCurrent(actionToken: number): boolean {
    return this.pendingInsert?.token === actionToken && this.actionGeneration === actionToken;
  }

  private finishAction(actionToken: number): void {
    if (this.isActionCurrent(actionToken)) this.pendingInsert = undefined;
  }

  private hostFailure(title: string): void {
    this.publish({
      phase: "blocked",
      title,
      detail: "Verseform could not confirm the result. Inspect the document before retrying or reopening the task pane.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  private cleanupFailure(title: string): void {
    this.publish({
      phase: "blocked",
      title,
      detail: "Verseform could not confirm the result. Inspect the document before retrying cleanup or reopening the task pane.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  private reconciliationFailure(): void {
    this.publish({
      phase: "blocked",
      title: "Verseform could not clear previous temporary annotations",
      detail: "No new annotations were added. Inspect the document, then reopen the task pane to retry cleanup.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  private ownershipFailure(): void {
    this.publish({
      phase: "blocked",
      title: "Verseform could not reconcile temporary annotation ownership",
      detail: "No new annotations were added. Reopen the task pane after local temporary-annotation storage is available.",
      canInsert: false,
      canCancel: false,
      focusTarget: "status",
    });
  }

  private serial(task: () => Promise<void>): Promise<void> {
    const next = this.operation.then(task, task);
    this.operation = next.catch(() => undefined);
    return next;
  }

  private publish(next: VerseformState): void {
    this.state = {
      ...next,
      catalogPhase: this.catalogPhase,
      translations: this.translations,
      ...(this.selectedTranslationId ? { selectedTranslationId: this.selectedTranslationId } : {}),
      canSelectTranslation: this.catalogPhase === "ready" && next.phase !== "inserting",
      canClearCache: this.catalogPhase !== "loading" && next.phase !== "inserting",
    };
    this.onState(this.state);
  }
}

/** Compatibility export for VFW-010 tests and historical callers. */
export { VerseformController as Vfw010Controller };
