import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import {
  BrowserAnnotationOwnership,
  MAX_MISSING_SWEEPS,
  MAX_OWNED_ANNOTATIONS,
  MemoryAnnotationOwnership,
  type AnnotationOwnership,
} from "../src/app/annotationOwnership";
import { Vfw010Controller, type Vfw010State } from "../src/app/controller";
import { candidatesForParagraph, type ParagraphSnapshot } from "../src/app/interaction";
import type { AnnotatedReference } from "../src/core/freshness";
import { scanReferences, type ReferenceCandidate } from "../src/core/reference";
import type { ReplacementResult, WordGateway, WordHostHandlers, WordRuntime } from "../src/office/gateway";
import { inspectWordHost, officeReadinessFailure } from "../src/office/capabilities";
import { isReplaceableAnnotationState } from "../src/office/replacementGuard";
import { OfficeWordGateway } from "../src/office/wordGateway";
import { mountTaskPane } from "../src/ui/taskPane";
import { taskPaneView } from "../src/ui/viewModel";
import { MemoryTranslationPreference } from "../src/adapters/translationPreference";
import type {
  Passage,
  ScriptureProvider,
  Translation,
  TranslationCatalog,
  TranslationPreferenceStore,
} from "../src/core/scripture";

type FakeParagraph = { text: string; revision: number };
type FakeAnnotation = { paragraphId: string; candidate: ReferenceCandidate };
type ReplacementHistory = { annotationId: string; annotation: FakeAnnotation; previousText: string };

class FakeWordGateway implements WordGateway {
  public readonly paragraphs = new Map<string, FakeParagraph>();
  public readonly annotations = new Map<string, FakeAnnotation>();
  public readonly handlerSnapshots: WordHostHandlers[] = [];
  public readonly removedBatches: string[][] = [];
  public readonly operations: string[] = [];
  public startCalls = 0;
  public stopCalls = 0;
  public replaceCalls = 0;
  public annotateCalls = 0;
  public failStart = false;
  public failRead = false;
  public failReplace = false;
  public readonly failingRemovalIds = new Set<string>();
  private readonly handlers = new Set<WordHostHandlers>();
  private nextAnnotation = 0;
  private readonly replacements: ReplacementHistory[] = [];
  private readBarrier: Promise<void> | undefined;
  private releaseReadBarrier: (() => void) | undefined;

  public constructor(paragraphId: string, text: string) {
    this.addParagraph(paragraphId, text);
  }

  public addParagraph(paragraphId: string, text: string): void {
    this.paragraphs.set(paragraphId, { text, revision: 0 });
  }

  public async start(handlers: WordHostHandlers): Promise<WordRuntime> {
    this.startCalls += 1;
    this.operations.push("start");
    this.handlerSnapshots.push(handlers);
    if (this.failStart) throw new Error("Fake start failure");
    this.handlers.add(handlers);
    let stopped = false;
    return {
      stop: async () => {
        if (stopped) return;
        stopped = true;
        this.stopCalls += 1;
        this.handlers.delete(handlers);
      },
    };
  }

  public async readParagraph(paragraphId: string): Promise<ParagraphSnapshot | undefined> {
    if (this.failRead) throw new Error("Fake read failure");
    await this.readBarrier;
    const paragraph = this.paragraphs.get(paragraphId);
    return paragraph ? { paragraphId, ...paragraph } : undefined;
  }

  public async annotate(paragraph: ParagraphSnapshot, candidate: ReferenceCandidate): Promise<string | undefined> {
    this.annotateCalls += 1;
    const current = this.paragraphs.get(paragraph.paragraphId);
    if (!current || current.text !== paragraph.text) return undefined;
    const annotationId = `annotation-${++this.nextAnnotation}`;
    this.operations.push(`annotate:${annotationId}`);
    this.annotations.set(annotationId, { paragraphId: paragraph.paragraphId, candidate });
    return annotationId;
  }

  public async removeAnnotations(annotationIds: readonly string[]): Promise<{
    removedAnnotationIds: string[];
    missingAnnotationIds: string[];
    failedAnnotationIds: string[];
  }> {
    this.removedBatches.push([...annotationIds]);
    this.operations.push(`remove:${annotationIds.join(",")}`);
    const removedAnnotationIds: string[] = [];
    const missingAnnotationIds: string[] = [];
    const failedAnnotationIds: string[] = [];
    for (const annotationId of annotationIds) {
      if (this.failingRemovalIds.has(annotationId)) {
        failedAnnotationIds.push(annotationId);
      } else if (this.annotations.delete(annotationId)) {
        removedAnnotationIds.push(annotationId);
      } else {
        missingAnnotationIds.push(annotationId);
      }
    }
    return { removedAnnotationIds, missingAnnotationIds, failedAnnotationIds };
  }

  public async replace(annotation: AnnotatedReference, replacementText: string): Promise<ReplacementResult> {
    this.replaceCalls += 1;
    if (this.failReplace) throw new Error("Fake replacement failure");
    const stored = this.annotations.get(annotation.annotationId);
    const paragraph = this.paragraphs.get(annotation.paragraphId);
    if (!stored || !paragraph) return "missing";
    if (stored.paragraphId !== annotation.paragraphId
      || stored.candidate.from !== annotation.range.from
      || stored.candidate.to !== annotation.range.to
      || paragraph.text !== annotation.paragraphText
      || paragraph.text.slice(annotation.range.from, annotation.range.to) !== annotation.sourceText) {
      return "stale";
    }
    this.replacements.push({
      annotationId: annotation.annotationId,
      annotation: stored,
      previousText: paragraph.text,
    });
    paragraph.text = `${paragraph.text.slice(0, annotation.range.from)}${replacementText}${paragraph.text.slice(annotation.range.to)}`;
    paragraph.revision += 1;
    this.annotations.delete(annotation.annotationId);
    return "replaced";
  }

  public async isAnnotationCurrent(annotation: AnnotatedReference): Promise<boolean> {
    const stored = this.annotations.get(annotation.annotationId);
    const paragraph = this.paragraphs.get(annotation.paragraphId);
    return Boolean(stored
      && paragraph
      && stored.paragraphId === annotation.paragraphId
      && stored.candidate.from === annotation.range.from
      && stored.candidate.to === annotation.range.to
      && paragraph.text === annotation.paragraphText
      && paragraph.text.slice(annotation.range.from, annotation.range.to) === annotation.sourceText);
  }

  public async undoLastReplacement(
    additionallyRestored: readonly [string, FakeAnnotation][] = [],
  ): Promise<void> {
    const replacement = this.replacements.pop();
    if (!replacement) throw new Error("No fake replacement to undo");
    const paragraph = this.paragraphs.get(replacement.annotation.paragraphId);
    if (!paragraph) throw new Error("Missing fake replacement paragraph");
    paragraph.text = replacement.previousText;
    paragraph.revision += 1;
    this.annotations.set(replacement.annotationId, replacement.annotation);
    for (const [annotationId, annotation] of additionallyRestored) {
      this.annotations.set(annotationId, annotation);
    }
    await this.emitChanged([replacement.annotation.paragraphId]);
  }

  public async change(paragraphId: string, text: string): Promise<void> {
    const paragraph = this.paragraphs.get(paragraphId);
    if (!paragraph) throw new Error("Missing fake paragraph");
    paragraph.text = text;
    paragraph.revision += 1;
    await this.emitChanged([paragraphId]);
  }

  public async emitChanged(paragraphIds: readonly string[]): Promise<void> {
    await Promise.all([...this.handlers].map(async (handler) => handler.onParagraphChanged(paragraphIds)));
  }

  public async mutateWithoutEvent(paragraphId: string, text: string): Promise<void> {
    const paragraph = this.paragraphs.get(paragraphId);
    if (!paragraph) throw new Error("Missing fake paragraph");
    paragraph.text = text;
    paragraph.revision += 1;
  }

  public async activate(annotationId: string, activation: "hovered" | "clicked" = "hovered"): Promise<void> {
    await Promise.all([...this.handlers].map(async (handler) => handler.onAnnotationActivated(annotationId, activation)));
  }

  public async removeFromWord(annotationId: string): Promise<void> {
    this.annotations.delete(annotationId);
    await Promise.all([...this.handlers].map(async (handler) => handler.onAnnotationRemoved([annotationId])));
  }

  public deferRead(): () => void {
    this.readBarrier = new Promise<void>((resolve) => {
      this.releaseReadBarrier = () => {
        this.readBarrier = undefined;
        this.releaseReadBarrier = undefined;
        resolve();
      };
    });
    return () => this.releaseReadBarrier?.();
  }

  /** Simulates an unloaded task-pane document where pagehide cleanup cannot run. */
  public abandonRuntime(): void {
    this.handlers.clear();
  }
}

class DeferredReserveOwnership extends MemoryAnnotationOwnership {
  private nextGate: { started: () => void; released: Promise<void> } | undefined;

  public deferNextReserve(): { started: Promise<void>; release: () => void } {
    let markStarted!: () => void;
    let release!: () => void;
    const started = new Promise<void>((resolve) => { markStarted = resolve; });
    const released = new Promise<void>((resolve) => { release = resolve; });
    this.nextGate = { started: markStarted, released };
    return { started, release };
  }

  public override async reserve(): Promise<string | undefined> {
    const gate = this.nextGate;
    this.nextGate = undefined;
    if (gate) {
      gate.started();
      await gate.released;
    }
    return super.reserve();
  }
}

/** Minimal exclusive-lock scheduler for two browser-store clients. */
class SerializedBrowserLocks {
  private tail = Promise.resolve();

  public request<T>(
    _name: string,
    _options: unknown,
    callback: () => T | Promise<T>,
  ): Promise<T> {
    const task = this.tail.then(callback, callback);
    this.tail = task.then(() => undefined, () => undefined);
    return task;
  }
}

function sharedBrowserStore(initial: unknown = []): {
  storage: { getItem: ReturnType<typeof vi.fn>; setItem: ReturnType<typeof vi.fn> };
  raw(): string;
} {
  let current = JSON.stringify(initial);
  return {
    storage: {
      getItem: vi.fn(() => current),
      setItem: vi.fn((_key: string, value: string) => { current = value; }),
    },
    raw: () => current,
  };
}

function controllerFor(
  word: FakeWordGateway,
  ownership: AnnotationOwnership = new MemoryAnnotationOwnership(),
  scripture?: ScriptureProvider,
  preference?: TranslationPreferenceStore,
): { controller: Vfw010Controller; states: Vfw010State[] } {
  const states: Vfw010State[] = [];
  const controller = new Vfw010Controller(
    word,
    (state) => states.push({ ...state }),
    ownership,
    scripture,
    preference,
  );
  return { controller, states };
}

describe("VFW-010 interaction kernel", () => {
  it("waits for a delimiter and uses UTF-16 offsets that include surrogate pairs", async () => {
    const match = scanReferences("🙂 John 3:16.")[0];
    expect(match).toMatchObject({ from: 3, to: 12, sourceText: "John 3:16" });

    const word = new FakeWordGateway("paragraph-1", "Read John 3:16");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16");
    expect(word.annotations.size).toBe(0);

    await word.change("paragraph-1", "Read John 3:16.");
    expect([...word.annotations.values()]).toMatchObject([{
      candidate: { sourceText: "John 3:16", from: 5, to: 14 },
    }]);
  });

  it("maps a hovered duplicate reference by annotation ID without a request, then replaces only that occurrence", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16 and John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "John 3:16 and John 3:16.");
    const ids = [...word.annotations.keys()];
    expect(ids).toHaveLength(2);

    const fetchSpy = vi.fn();
    vi.stubGlobal("fetch", fetchSpy);
    try {
      await word.activate(ids[1]!, "hovered");
      expect(controller.getState()).toMatchObject({
        phase: "preview",
        selectedAnnotationId: ids[1],
        canInsert: true,
        canCancel: true,
        preview: { heading: "John 3:16" },
      });
      expect(fetchSpy).not.toHaveBeenCalled();
      await controller.insertSelected();
    } finally {
      vi.unstubAllGlobals();
    }

    const output = word.paragraphs.get("paragraph-1")?.text ?? "";
    expect(output).toBe(
      "John 3:16 and [Verseform VFW-010 local fake passage. No Scripture text is bundled.] (John 3:16, VFW-010 test text).",
    );
    expect(candidatesForParagraph(output)).toMatchObject([{ from: 0, sourceText: "John 3:16" }]);
    expect(candidatesForParagraph(output)).toHaveLength(1);
    expect(word.replaceCalls).toBe(1);
  });

  it("excludes an inserted NASB citation while continuing to detect ordinary references", () => {
    const output = "John 1:1; text (John 3:16, NASB) and John 3:17.";
    expect(candidatesForParagraph(output).map((candidate) => candidate.sourceText)).toEqual([
      "John 1:1",
      "John 3:17",
    ]);
  });

  it("refuses stale text after preview and clears its host annotation without replacing", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    await word.mutateWithoutEvent("paragraph-1", "Please read John 3:16.");

    await expect(controller.insertSelected()).resolves.toBeUndefined();
    expect(controller.getState()).toMatchObject({
      phase: "watching",
      title: "Writing changed before insertion",
      canInsert: false,
      focusTarget: "status",
    });
    expect(word.annotations.has(annotationId!)).toBe(false);
    expect(word.replaceCalls).toBe(0);
  });

  it("allows only one immediate Insert action while a fresh-read check is pending", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    const releaseRead = word.deferRead();

    const first = controller.insertSelected();
    const second = controller.insertSelected();
    expect(controller.getState()).toMatchObject({ phase: "inserting", canInsert: false, canCancel: true });
    releaseRead();
    await Promise.all([first, second]);

    expect(word.replaceCalls).toBe(1);
    expect(controller.getState()).toMatchObject({ phase: "complete", title: "VFW-010 test replacement inserted" });
  });

  it("lets Clear cancel a pending fresh-read check before any mutation boundary", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    const releaseRead = word.deferRead();

    const insert = controller.insertSelected();
    await controller.clearSelection();
    releaseRead();
    await insert;

    expect(word.replaceCalls).toBe(0);
    expect(controller.getState()).toMatchObject({
      phase: "watching",
      title: "Reference selection cleared",
      canInsert: false,
    });
    expect(word.paragraphs.get("paragraph-1")?.text).toBe("Read John 3:16.");
  });

  it("cancels a deferred ownership reservation when the runtime stops before annotation creation", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const ownership = new DeferredReserveOwnership();
    const { controller } = controllerFor(word, ownership);
    await controller.start();
    const gate = ownership.deferNextReserve();

    const refresh = word.change("paragraph-1", "Read John 3:16.");
    await gate.started;
    await controller.stop();
    gate.release();
    await refresh;

    expect(word.annotateCalls).toBe(0);
    expect(word.annotations.size).toBe(0);
    expect(await ownership.load()).toEqual([]);
  });

  it("refuses an unknown or host-removed selected annotation and returns focus to status", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();

    await word.activate("not-ours");
    expect(controller.getState()).toMatchObject({ phase: "watching", canInsert: false, focusTarget: "status" });

    await word.activate(annotationId!);
    await word.removeFromWord(annotationId!);
    expect(controller.getState()).toMatchObject({
      title: "Reference annotation was removed",
      canInsert: false,
      focusTarget: "status",
    });
  });

  it("serializes multi-ID and burst paragraph events without making a request", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    word.addParagraph("paragraph-2", "Read Romans 8:28.");
    const { controller } = controllerFor(word);
    await controller.start();

    const fetchSpy = vi.fn();
    vi.stubGlobal("fetch", fetchSpy);
    try {
      await Promise.all([
        word.emitChanged(["paragraph-1", "paragraph-2"]),
        word.emitChanged(["paragraph-1"]),
      ]);
    } finally {
      vi.unstubAllGlobals();
    }
    expect(fetchSpy).not.toHaveBeenCalled();
    expect([...word.annotations.values()].filter((annotation) => annotation.paragraphId === "paragraph-1")).toHaveLength(1);
    expect([...word.annotations.values()].filter((annotation) => annotation.paragraphId === "paragraph-2")).toHaveLength(1);
  });

  it("fails closed after a partial old-annotation cleanup failure and retries on a later change", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16 and Romans 8:28.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "John 3:16 and Romans 8:28.");
    const oldIds = [...word.annotations.keys()];
    expect(oldIds).toHaveLength(2);
    word.failingRemovalIds.add(oldIds[0]!);

    await word.change("paragraph-1", "Read John 3:16.");
    expect(controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Verseform could not clear previous annotations",
      canInsert: false,
    });
    expect(word.annotations.has(oldIds[0]!)).toBe(true);
    expect(word.annotations.size).toBe(1);

    word.failingRemovalIds.clear();
    await word.change("paragraph-1", "Read John 3:16.");
    expect(controller.getState()).toMatchObject({ phase: "watching", title: "Reference ready" });
    expect(word.annotations.size).toBe(1);
    expect([...word.annotations.keys()][0]).not.toBe(oldIds[0]);
  });

  it("reports a stale cleanup failure instead of claiming the annotation was cleared", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    word.failingRemovalIds.add(annotationId!);
    await word.mutateWithoutEvent("paragraph-1", "Please read John 3:16.");

    await controller.insertSelected();
    expect(controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Verseform could not clear the stale annotation",
    });
    expect(word.annotations.has(annotationId!)).toBe(true);

    word.failingRemovalIds.clear();
    await word.change("paragraph-1", "Please read John 3:16.");
    expect(controller.getState()).toMatchObject({ phase: "watching", title: "Reference ready" });
    expect(word.annotations.size).toBe(1);
    expect([...word.annotations.keys()]).not.toContain(annotationId);
  });

  it("contains start, read, replacement, and removal failures without unhandled action rejection", async () => {
    const failedStart = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    failedStart.failStart = true;
    const startAttempt = controllerFor(failedStart);
    await expect(startAttempt.controller.start()).resolves.toBeUndefined();
    expect(startAttempt.controller.getState()).toMatchObject({ phase: "blocked", title: "Word events could not start" });

    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    word.failRead = true;
    await expect(word.emitChanged(["paragraph-1"])).resolves.toBeUndefined();
    expect(controller.getState()).toMatchObject({ phase: "blocked", title: "Word could not refresh this reference" });

    const replacementWord = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const replacement = controllerFor(replacementWord);
    await replacement.controller.start();
    await replacementWord.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = replacementWord.annotations.keys();
    await replacementWord.activate(annotationId!);
    replacementWord.failReplace = true;
    await expect(replacement.controller.insertSelected()).resolves.toBeUndefined();
    expect(replacement.controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Word did not confirm the replacement",
      detail: expect.stringMatching(/could not confirm the result\. Inspect the document/u),
    });
    expect(replacementWord.annotations.has(annotationId!)).toBe(false);

    const cleanupWord = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const cleanup = controllerFor(cleanupWord);
    await cleanup.controller.start();
    await cleanupWord.change("paragraph-1", "Read John 3:16.");
    const [cleanupAnnotationId] = cleanupWord.annotations.keys();
    await cleanupWord.activate(cleanupAnnotationId!);
    cleanupWord.failReplace = true;
    cleanupWord.failingRemovalIds.add(cleanupAnnotationId!);
    await expect(cleanup.controller.insertSelected()).resolves.toBeUndefined();
    expect(cleanup.controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Verseform could not clear the failed annotation",
      detail: expect.stringMatching(/could not confirm the result\. Inspect the document/u),
    });
    expect(cleanupWord.annotations.has(cleanupAnnotationId!)).toBe(true);
  });

  it("starts and stops idempotently, removes handlers, and ignores delayed additive-handler events", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const first = controllerFor(word);
    await Promise.all([first.controller.start(), first.controller.start()]);
    expect(word.startCalls).toBe(1);
    await word.change("paragraph-1", "Read John 3:16.");
    const staleHandlers = [...word.handlerSnapshots];
    await Promise.all([first.controller.stop(), first.controller.stop()]);
    expect(word.stopCalls).toBe(1);
    expect(word.annotations.size).toBe(0);
    await staleHandlers[0]!.onParagraphChanged(["paragraph-1"]);
    expect(word.annotations.size).toBe(0);

    const reopened = controllerFor(word);
    await reopened.controller.start();
    await word.emitChanged(["paragraph-1"]);
    expect(word.startCalls).toBe(2);
    expect(word.annotations.size).toBe(1);
  });

  it("revalidates an annotation resurrected by Word Undo instead of layering a new marker", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    const { controller } = controllerFor(word, ownership);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    await controller.insertSelected();

    expect(word.annotations.size).toBe(0);
    await word.undoLastReplacement();

    expect(word.annotations.size).toBe(1);
    expect([...word.annotations.keys()]).toEqual([annotationId]);
    expect(word.annotateCalls).toBe(1);
    expect(await ownership.load()).toEqual([{ annotationId, missingSweeps: 0 }]);
    await word.activate(annotationId!);
    expect(controller.getState()).toMatchObject({
      phase: "preview",
      title: "Preview ready",
      selectedAnnotationId: annotationId,
      canInsert: true,
    });
  });

  it("keeps a replaced annotation ID owned when stop cleans its retired runtime record", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    const { controller } = controllerFor(word, ownership);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    await controller.insertSelected();

    await controller.stop();

    expect(word.removedBatches.at(-1)).toEqual([annotationId]);
    expect(await ownership.load()).toEqual([{ annotationId, missingSweeps: 0 }]);
  });

  it("restores every exact annotation from a refreshed paragraph when Undo resurrects Word's prior marker set", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 1:1 and John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    const { controller } = controllerFor(word, ownership);
    await controller.start();
    await word.change("paragraph-1", "John 1:1 and John 3:16.");
    const original = [...word.annotations.entries()];
    const [firstAnnotationId, firstAnnotation] = original[0]!;
    const [selectedAnnotationId] = original[1]!;
    await word.activate(selectedAnnotationId);
    await controller.insertSelected();

    // Model Word's annotation-removed event arriving before the paragraph
    // refresh that marks the still-valid John 1:1 in the changed paragraph.
    await word.removeFromWord(firstAnnotationId);
    await word.emitChanged(["paragraph-1"]);
    expect(word.annotations.size).toBe(1);
    expect(word.annotateCalls).toBe(3);

    // Word Undo restores the full pre-replacement paragraph and the native
    // marker set that existed before the replacement transaction.
    await word.undoLastReplacement([[firstAnnotationId, firstAnnotation]]);

    expect([...word.annotations.keys()].sort()).toEqual([firstAnnotationId, selectedAnnotationId].sort());
    expect(word.annotations.size).toBe(2);
    expect(word.annotateCalls).toBe(3);
    await word.activate(firstAnnotationId);
    expect(controller.getState()).toMatchObject({
      phase: "preview",
      title: "Preview ready",
      selectedAnnotationId: firstAnnotationId,
      canInsert: true,
    });
  });

  it("revalidates a retired ID on activation when Word resurrects it after refresh already marked the range again", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [retiredId, retiredAnnotation] = [...word.annotations.entries()][0]!;

    // The first event removes the old native marker. The subsequent paragraph
    // refresh sees no marker and creates a new current one before Word later
    // restores the original ID from its Undo history.
    await word.removeFromWord(retiredId);
    await word.emitChanged(["paragraph-1"]);
    expect(word.annotations.size).toBe(1);
    expect([...word.annotations.keys()]).not.toContain(retiredId);
    word.annotations.set(retiredId, retiredAnnotation);

    await word.activate(retiredId);
    expect([...word.annotations.keys()]).toEqual([retiredId]);
    expect(word.annotations.size).toBe(1);
    expect(controller.getState()).toMatchObject({
      phase: "preview",
      title: "Preview ready",
      selectedAnnotationId: retiredId,
      canInsert: true,
    });
  });

  it("sweeps exact persisted ownership before a reopened pane binds new handlers", async () => {
    const word = new FakeWordGateway("paragraph-1", "🙂 John 3:16 and John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    const first = controllerFor(word, ownership);
    await first.controller.start();
    await word.change("paragraph-1", "🙂 John 3:16 and John 3:16.");
    const staleIds = [...word.annotations.keys()];
    expect(staleIds).toHaveLength(2);
    expect(await ownership.load()).toEqual(staleIds.map((annotationId) => ({ annotationId, missingSweeps: 0 })));

    // A native task-pane close can end pagehide before its async stop work
    // reaches Word. Model that process loss without giving the old controller
    // a chance to clear its annotations.
    word.abandonRuntime();
    const reopened = controllerFor(word, ownership);
    await reopened.controller.start();

    expect(word.annotations.size).toBe(0);
    expect(await ownership.load()).toEqual(staleIds.map((annotationId) => ({ annotationId, missingSweeps: 0 })));
    expect(word.removedBatches.at(-1)).toEqual(staleIds);
    const sweep = word.operations.lastIndexOf(`remove:${staleIds.join(",")}`);
    expect(sweep).toBeGreaterThan(-1);
    expect(word.operations.slice(sweep + 1)).toContain("start");

    await word.change("paragraph-1", "🙂 John 3:16 and John 3:16. Again John 1:1.");
    expect(word.annotations.size).toBe(3);
    expect([...word.annotations.keys()]).not.toEqual(expect.arrayContaining(staleIds));
  });

  it("fails closed and retains exact ownership when startup reconciliation cannot clear it", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    const first = controllerFor(word, ownership);
    await first.controller.start();
    await word.change("paragraph-1", "John 3:16.");
    const [staleId] = word.annotations.keys();
    word.failingRemovalIds.add(staleId!);
    word.abandonRuntime();

    const blocked = controllerFor(word, ownership);
    await blocked.controller.start();
    expect(blocked.controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Verseform could not clear previous temporary annotations",
    });
    expect(word.startCalls).toBe(1);
    expect(word.annotations.has(staleId!)).toBe(true);
    expect(await ownership.load()).toEqual([{ annotationId: staleId, missingSweeps: 0 }]);

    word.failingRemovalIds.clear();
    const recovered = controllerFor(word, ownership);
    await recovered.controller.start();
    expect(recovered.controller.getState()).toMatchObject({ phase: "watching" });
    expect(word.annotations.has(staleId!)).toBe(false);
    expect(await ownership.load()).toEqual([{ annotationId: staleId, missingSweeps: 0 }]);
    expect(word.startCalls).toBe(2);
  });

  it("blocks before Word starts when exact annotation ownership cannot be stored", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16.");
    const unavailableOwnership: AnnotationOwnership = {
      load: async () => { throw new Error("storage unavailable"); },
      reserve: async () => undefined,
      commit: async () => false,
      cancel: async () => undefined,
      release: async () => undefined,
      recordRemoval: async () => undefined,
    };
    const { controller } = controllerFor(word, unavailableOwnership);

    await controller.start();
    expect(controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Verseform could not reconcile temporary annotation ownership",
    });
    expect(word.startCalls).toBe(0);
    expect(word.annotations.size).toBe(0);
  });

  it("does not age a cross-document missing ID again when that pane stops", async () => {
    const word = new FakeWordGateway("paragraph-1", "No reference.");
    const ownership = new MemoryAnnotationOwnership();
    await ownership.seed([{ annotationId: "orphan-from-another-document", missingSweeps: 0 }]);
    const { controller } = controllerFor(word, ownership);

    await controller.start();
    await controller.stop();

    expect(await ownership.load()).toEqual([{ annotationId: "orphan-from-another-document", missingSweeps: 1 }]);
    expect(word.removedBatches).toEqual([["orphan-from-another-document"]]);
  });

  it("fails closed without overwriting a malformed browser ownership record", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16.");
    const setItem = vi.fn();
    vi.stubGlobal("localStorage", {
      getItem: vi.fn(() => '[{"annotationId":"orphan","missingSweeps":0},"malformed"]'),
      setItem,
    });
    vi.stubGlobal("navigator", {
      locks: { request: async (_name: string, _options: unknown, callback: () => Promise<unknown>) => callback() },
    });
    try {
      const { controller } = controllerFor(word, new BrowserAnnotationOwnership());
      await controller.start();

      expect(controller.getState()).toMatchObject({
        phase: "blocked",
        title: "Verseform could not reconcile temporary annotation ownership",
      });
      expect(setItem).not.toHaveBeenCalled();
      expect(word.startCalls).toBe(0);
      expect(word.annotations.size).toBe(0);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("blocks before Word event binding when browser ownership locking is unavailable", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16.");
    vi.stubGlobal("localStorage", { getItem: vi.fn(() => null), setItem: vi.fn() });
    vi.stubGlobal("navigator", {});
    try {
      const { controller } = controllerFor(word, new BrowserAnnotationOwnership());
      await controller.start();

      expect(controller.getState()).toMatchObject({
        phase: "blocked",
        title: "Verseform could not reconcile temporary annotation ownership",
      });
      expect(word.startCalls).toBe(0);
      expect(word.annotations.size).toBe(0);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("uses browser locking to reread shared ownership for concurrent claims, deltas, and the final slot", async () => {
    const store = sharedBrowserStore();
    const locks = new SerializedBrowserLocks();
    vi.stubGlobal("localStorage", store.storage);
    vi.stubGlobal("navigator", { locks });
    try {
      const first = new BrowserAnnotationOwnership();
      const second = new BrowserAnnotationOwnership();
      const [reservationA, reservationB] = await Promise.all([first.reserve(), second.reserve()]);
      expect(reservationA).toBeTruthy();
      expect(reservationB).toBeTruthy();
      await Promise.all([
        first.commit(reservationA!, "browser-annotation-a"),
        second.commit(reservationB!, "browser-annotation-b"),
      ]);
      expect(await first.load()).toEqual([
        { annotationId: "browser-annotation-a", missingSweeps: 0 },
        { annotationId: "browser-annotation-b", missingSweeps: 0 },
      ]);

      await Promise.all([
        first.recordRemoval({ removedAnnotationIds: [], missingAnnotationIds: ["browser-annotation-a"] }),
        second.recordRemoval({ removedAnnotationIds: ["browser-annotation-b"], missingAnnotationIds: [] }),
      ]);
      expect(await second.load()).toEqual([{ annotationId: "browser-annotation-a", missingSweeps: 1 }]);

      const finalStore = sharedBrowserStore(Array.from({ length: MAX_OWNED_ANNOTATIONS - 1 }, (_, index) => ({
        annotationId: `existing-${index}`,
        missingSweeps: 0,
      })));
      vi.stubGlobal("localStorage", finalStore.storage);
      const [lastSlotA, lastSlotB] = await Promise.all([
        new BrowserAnnotationOwnership().reserve(),
        new BrowserAnnotationOwnership().reserve(),
      ]);
      expect([lastSlotA, lastSlotB].filter(Boolean)).toHaveLength(1);
      expect(JSON.parse(finalStore.raw())).toHaveLength(MAX_OWNED_ANNOTATIONS);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("does not leave an unrecorded marker when the bounded ownership ledger is full", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    await ownership.seed(Array.from({ length: MAX_OWNED_ANNOTATIONS }, (_, index) => ({
      annotationId: `old-document-${index}`,
      missingSweeps: 0,
    })));
    const { controller } = controllerFor(word, ownership);

    await controller.start();
    await word.change("paragraph-1", "John 3:16.");
    expect(controller.getState()).toMatchObject({
      phase: "blocked",
      title: "Verseform could not reserve temporary annotation ownership",
    });
    expect(word.annotations.size).toBe(0);
    expect((await ownership.load())).toHaveLength(MAX_OWNED_ANNOTATIONS);
  });

  it("retains an orphaned ID through another document's missing sweep, then recovers it in its owner document", async () => {
    const word = new FakeWordGateway("paragraph-1", "John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    const documentA = controllerFor(word, ownership);
    await documentA.controller.start();
    await word.change("paragraph-1", "John 3:16.");
    const [orphanId, orphan] = [...word.annotations.entries()][0]!;
    word.abandonRuntime();

    // Reuse the fake port as document B: A's orphan exists only in document A.
    word.annotations.clear();
    const documentB = controllerFor(word, ownership);
    await documentB.controller.start();
    expect(documentB.controller.getState()).toMatchObject({ phase: "watching" });
    expect(await ownership.load()).toEqual([{ annotationId: orphanId, missingSweeps: 1 }]);
    word.abandonRuntime();

    // Reopen document A before the bounded missing-ID expiry.
    word.annotations.set(orphanId, orphan);
    const reopenedA = controllerFor(word, ownership);
    await reopenedA.controller.start();
    expect(reopenedA.controller.getState()).toMatchObject({ phase: "watching" });
    expect(word.annotations.has(orphanId)).toBe(false);
    expect(await ownership.load()).toEqual([{ annotationId: orphanId, missingSweeps: 1 }]);
  });

  it("distinguishes a late post-close resurrection by its opaque prior ID and refuses to guess its range", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const ownership = new MemoryAnnotationOwnership();
    await ownership.seed([{ annotationId: "previous-pane-id", missingSweeps: 0 }]);
    const { controller } = controllerFor(word, ownership);
    await controller.start();

    // Startup could not find this ID. Model Word Undo restoring it later,
    // after the former runtime's paragraph/range metadata was forgotten.
    word.annotations.set("previous-pane-id", {
      paragraphId: "paragraph-1",
      candidate: candidatesForParagraph("Read John 3:16.")[0]!,
    });
    await word.activate("previous-pane-id");

    expect(controller.getState()).toMatchObject({
      phase: "watching",
      title: "A previous temporary annotation returned",
      detail: expect.stringMatching(/cannot safely map it after reopening/u),
      canInsert: false,
      focusTarget: "status",
    });
    expect(word.annotations.has("previous-pane-id")).toBe(true);
    expect(word.replaceCalls).toBe(0);
  });

  it("bounds repeated missing-ID retries without blocking a new document", async () => {
    const word = new FakeWordGateway("paragraph-1", "No reference.");
    const ownership = new MemoryAnnotationOwnership();
    await ownership.seed([{ annotationId: "orphan-from-another-document", missingSweeps: 0 }]);

    for (let attempt = 0; attempt < MAX_MISSING_SWEEPS; attempt += 1) {
      const { controller } = controllerFor(word, ownership);
      await controller.start();
      expect(controller.getState()).toMatchObject({ phase: "watching" });
      word.abandonRuntime();
    }

    expect(await ownership.load()).toEqual([]);
    expect(word.startCalls).toBe(MAX_MISSING_SWEEPS);
  });

  it("serializes parallel ownership claims and interleaved missing/removal deltas", async () => {
    const ownership = new MemoryAnnotationOwnership();
    const [reservationA, reservationB] = await Promise.all([ownership.reserve(), ownership.reserve()]);
    expect(reservationA).toBeTruthy();
    expect(reservationB).toBeTruthy();

    await Promise.all([
      ownership.commit(reservationA!, "annotation-a"),
      ownership.commit(reservationB!, "annotation-b"),
    ]);
    expect(await ownership.load()).toEqual([
      { annotationId: "annotation-a", missingSweeps: 0 },
      { annotationId: "annotation-b", missingSweeps: 0 },
    ]);

    await Promise.all([
      ownership.recordRemoval({ removedAnnotationIds: [], missingAnnotationIds: ["annotation-a"] }),
      ownership.recordRemoval({ removedAnnotationIds: [], missingAnnotationIds: ["annotation-b"] }),
    ]);
    expect(await ownership.load()).toEqual([
      { annotationId: "annotation-a", missingSweeps: 1 },
      { annotationId: "annotation-b", missingSweeps: 1 },
    ]);

    await Promise.all([
      ownership.recordRemoval({ removedAnnotationIds: ["annotation-a"], missingAnnotationIds: [] }),
      ownership.recordRemoval({ removedAnnotationIds: [], missingAnnotationIds: ["annotation-b"] }),
    ]);
    expect(await ownership.load()).toEqual([{ annotationId: "annotation-b", missingSweeps: 2 }]);
  });

  it("atomically reserves the final ledger slot before either concurrent host create can begin", async () => {
    const ownership = new MemoryAnnotationOwnership();
    await ownership.seed(Array.from({ length: MAX_OWNED_ANNOTATIONS - 1 }, (_, index) => ({
      annotationId: `old-document-${index}`,
      missingSweeps: 0,
    })));

    const [first, second] = await Promise.all([ownership.reserve(), ownership.reserve()]);
    expect([first, second].filter(Boolean)).toHaveLength(1);
    await ownership.cancel(first ?? second!);
    expect((await ownership.load())).toHaveLength(MAX_OWNED_ANNOTATIONS - 1);
  });

  it("offers a cancel path without changing Word text", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const { controller } = controllerFor(word);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    await controller.clearSelection();
    expect(controller.getState()).toMatchObject({ title: "Reference selection cleared", canInsert: false, canCancel: false });
    expect(word.paragraphs.get("paragraph-1")?.text).toBe("Read John 3:16.");
    expect(word.annotations.has(annotationId!)).toBe(true);
  });
});

const dbsTranslations: Translation[] = [
  {
    id: "ENGTEST",
    citationLabel: "TEST",
    name: "DBS Test Bible",
    attribution: "DBS Test Bible (TEST): DBS fixture copyright.",
  },
  {
    id: "ENGNASB",
    citationLabel: "NASB",
    name: "New American Standard Bible",
    attribution: "New American Standard Bible (NASB): © The Lockman Foundation",
  },
];

class RecordedScriptureProvider implements ScriptureProvider {
  public readonly calls: Array<{
    reference: Parameters<ScriptureProvider["getPassage"]>[0];
    translationId: string;
    signal?: AbortSignal;
  }> = [];
  public clearCalls = 0;
  public failCatalog = false;
  public failPassage = false;
  private deferred: {
    started: () => void;
    passage: Promise<void>;
  } | undefined;

  public async listTranslations(): Promise<TranslationCatalog> {
    if (this.failCatalog) throw new Error("DBS catalog fixture unavailable.");
    return { translations: dbsTranslations, cached: false };
  }

  public async getPassage(
    reference: Parameters<ScriptureProvider["getPassage"]>[0],
    translationId: string,
    signal?: AbortSignal,
  ): Promise<Passage> {
    this.calls.push({ reference: { ...reference }, translationId, signal });
    const deferred = this.deferred;
    this.deferred = undefined;
    if (deferred) {
      deferred.started();
      await deferred.passage;
    }
    if (this.failPassage) throw new Error("DBS fixture unavailable.");
    const translation = dbsTranslations.find((item) => item.id === translationId);
    if (!translation) throw new Error("Unknown test translation.");
    return {
      reference,
      display: `${reference.bookName} ${reference.chapter}:${reference.verseStart}${
        reference.verseEnd === undefined ? "" : `-${reference.verseEnd}`
      }`,
      translationId,
      citationLabel: translation.citationLabel,
      translationName: translation.name,
      attribution: translation.attribution,
      text: `DBS test verse ${reference.verseStart} for ${reference.bookId}.`,
      cached: false,
    };
  }

  public async clearCache(): Promise<void> { this.clearCalls += 1; }

  public deferNextPassage(): { started: Promise<void>; release: () => void } {
    let markStarted!: () => void;
    let release!: () => void;
    const started = new Promise<void>((resolve) => { markStarted = resolve; });
    const passage = new Promise<void>((resolve) => { release = resolve; });
    this.deferred = { started: markStarted, passage };
    return { started, release };
  }
}

describe("VFW-020 provider integration", () => {
  it("defaults to NASB, sends coordinates only, and inserts passage/citation/attribution in one replacement", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    expect(controller.getState()).toMatchObject({
      catalogPhase: "ready",
      selectedTranslationId: "ENGNASB",
      canSelectTranslation: true,
    });

    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    expect(controller.getState()).toMatchObject({
      phase: "preview",
      title: "Preview ready",
      preview: {
        heading: "John 3:16 · NASB",
        text: "DBS test verse 16 for JHN.",
        attribution: "New American Standard Bible (NASB): © The Lockman Foundation",
        translationId: "ENGNASB",
      },
    });
    expect(scripture.calls).toMatchObject([{
      reference: { bookId: "JHN", bookName: "John", chapter: 3, verseStart: 16 },
      translationId: "ENGNASB",
    }]);
    expect(JSON.stringify(scripture.calls[0]?.reference)).not.toContain("Read John");

    await controller.insertSelected();
    expect(word.paragraphs.get("paragraph-1")?.text).toBe(
      "Read DBS test verse 16 for JHN. (John 3:16, NASB)\n"
      + "New American Standard Bible (NASB): © The Lockman Foundation.",
    );
    expect(word.replaceCalls).toBe(1);
  });

  it("shares a pending hover preview with click and inserts the exact reference once", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    const deferred = scripture.deferNextPassage();

    const hover = word.activate(annotationId!, "hovered");
    await deferred.started;
    const click = word.activate(annotationId!, "clicked");
    deferred.release();
    await Promise.all([hover, click]);

    expect(scripture.calls).toHaveLength(1);
    expect(word.replaceCalls).toBe(1);
    expect(word.paragraphs.get("paragraph-1")?.text).toBe(
      "Read DBS test verse 16 for JHN. (John 3:16, NASB)\n"
      + "New American Standard Bible (NASB): © The Lockman Foundation.",
    );
  });

  it("inserts a hover-ready passage on click without requesting it again", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();

    await word.activate(annotationId!, "hovered");
    expect(scripture.calls).toHaveLength(1);
    expect(word.replaceCalls).toBe(0);

    await word.activate(annotationId!, "clicked");

    expect(scripture.calls).toHaveLength(1);
    expect(word.replaceCalls).toBe(1);
    expect(word.paragraphs.get("paragraph-1")?.text).toContain("(John 3:16, NASB)");
  });

  it("loads and inserts from one explicit annotation click without requiring hover", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();

    await word.activate(annotationId!, "clicked");

    expect(scripture.calls).toHaveLength(1);
    expect(word.replaceCalls).toBe(1);
    expect(word.paragraphs.get("paragraph-1")?.text).toContain("(John 3:16, NASB)");
  });

  it("refuses click-to-insert when the reference changes during its DBS request", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    const deferred = scripture.deferNextPassage();

    const click = word.activate(annotationId!, "clicked");
    await deferred.started;
    await word.mutateWithoutEvent("paragraph-1", "Read John 3:15.");
    deferred.release();
    await click;

    expect(controller.getState()).toMatchObject({
      phase: "watching",
      title: "Writing changed before insertion",
      canInsert: false,
    });
    expect(word.replaceCalls).toBe(0);
    expect(word.paragraphs.get("paragraph-1")?.text).toBe("Read John 3:15.");
  });

  it("honors a saved authorized translation and remembers an explicit change", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const preference = new MemoryTranslationPreference("ENGTEST");
    const { controller } = controllerFor(
      word,
      new MemoryAnnotationOwnership(),
      scripture,
      preference,
    );
    await controller.start();
    expect(controller.getState().selectedTranslationId).toBe("ENGTEST");
    await controller.selectTranslation("ENGNASB");
    await expect(preference.load()).resolves.toBe("ENGNASB");
  });

  it("rejects a passage result when Word changed while DBS was responding", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    const deferred = scripture.deferNextPassage();
    const activation = word.activate(annotationId!);
    await deferred.started;
    await word.mutateWithoutEvent("paragraph-1", "Read John 3:15.");
    deferred.release();
    await activation;

    expect(controller.getState()).toMatchObject({
      phase: "watching",
      title: "Writing changed before insertion",
      canInsert: false,
    });
    expect(word.replaceCalls).toBe(0);
    expect(word.paragraphs.get("paragraph-1")?.text).toBe("Read John 3:15.");
  });

  it("aborts and ignores an old preview when the translation changes", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    const deferred = scripture.deferNextPassage();
    const activation = word.activate(annotationId!);
    await deferred.started;
    const requestSignal = scripture.calls[0]?.signal;
    await controller.selectTranslation("ENGTEST");
    expect(requestSignal?.aborted).toBe(true);
    deferred.release();
    await activation;

    expect(controller.getState()).toMatchObject({
      phase: "watching",
      selectedTranslationId: "ENGTEST",
      canInsert: false,
    });
    expect(controller.getState().preview).toBeUndefined();
  });

  it("aborts an in-flight preview before clearing local Scripture content", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    const deferred = scripture.deferNextPassage();
    const activation = word.activate(annotationId!);
    await deferred.started;

    await controller.clearScriptureCache();
    expect(scripture.calls[0]?.signal?.aborted).toBe(true);
    expect(scripture.clearCalls).toBe(1);
    expect(controller.getState()).toMatchObject({
      title: "Local Scripture cache cleared",
      canInsert: false,
    });
    deferred.release();
    await activation;
    expect(controller.getState().preview).toBeUndefined();
  });

  it("keeps prose and detection intact when DBS fails and exposes cache clearing", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    scripture.failPassage = true;
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    await word.change("paragraph-1", "Read John 3:16.");
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);

    expect(controller.getState()).toMatchObject({
      title: "Scripture text is unavailable",
      canInsert: false,
    });
    expect(word.paragraphs.get("paragraph-1")?.text).toBe("Read John 3:16.");
    expect(word.annotations.has(annotationId!)).toBe(true);
    await controller.clearScriptureCache();
    expect(scripture.clearCalls).toBe(1);
    expect(controller.getState().title).toBe("Local Scripture cache cleared");
  });

  it("starts local Word detection even when the translation catalog is offline", async () => {
    const word = new FakeWordGateway("paragraph-1", "Read John 3:16.");
    const scripture = new RecordedScriptureProvider();
    scripture.failCatalog = true;
    const { controller } = controllerFor(word, new MemoryAnnotationOwnership(), scripture);
    await controller.start();
    expect(controller.getState()).toMatchObject({
      catalogPhase: "unavailable",
      title: "Scripture service is unavailable",
    });

    await word.change("paragraph-1", "Read John 3:16.");
    expect(word.annotations.size).toBe(1);
    const [annotationId] = word.annotations.keys();
    await word.activate(annotationId!);
    expect(controller.getState()).toMatchObject({
      title: "Scripture text needs a connection",
      canInsert: false,
    });
    expect(scripture.calls).toHaveLength(0);
  });
});

describe("replacement and task-pane contracts", () => {
  it("only permits Created annotations to cross the Word replacement boundary", () => {
    expect(isReplaceableAnnotationState("Created")).toBe(true);
    expect(isReplaceableAnnotationState("Accepted")).toBe(false);
    expect(isReplaceableAnnotationState("Rejected")).toBe(false);
    expect(isReplaceableAnnotationState(undefined)).toBe(false);
  });

  it("keeps explicit controls disabled until an annotation is fresh", () => {
    expect(taskPaneView({
      phase: "watching",
      title: "Watching",
      detail: "Local only.",
      canInsert: false,
      canCancel: false,
    })).toMatchObject({
      statusKind: "ready",
      preview: { hidden: true },
      insert: { disabled: true, label: "Insert passage" },
      cancel: { disabled: true },
    });
    expect(taskPaneView({
      phase: "preview",
      title: "Preview ready",
      detail: "No request.",
      canInsert: true,
      canCancel: true,
      preview: {
        heading: "John 3:16", text: "Fake test data.", attribution: "Local test data.",
        insertText: "Fake test data.", translationId: "VFW-010-FAKE",
        translationName: "VFW-010 test text", citationLabel: "VFW-010 test text", cached: false,
      },
    })).toMatchObject({
      preview: { hidden: false, heading: "John 3:16" },
      insert: { disabled: false, label: "Insert test passage" },
      cancel: { disabled: false },
    });
  });
});

describe("Word host capability contract", () => {
  it("provides clear no-change states for a non-Word host, unsupported WordApi, and readiness failure", () => {
    const supported = vi.fn(() => true);
    vi.stubGlobal("Office", {
      HostType: { Word: "Word" },
      context: { host: "Excel", requirements: { isSetSupported: supported } },
    });
    try {
      expect(inspectWordHost()).toEqual({
        kind: "blocked",
        title: "Open in Word",
        detail: "This add-in works only inside Microsoft Word.",
      });
      expect(supported).not.toHaveBeenCalled();

      vi.stubGlobal("Office", {
        HostType: { Word: "Word" },
        context: { host: "Word", requirements: { isSetSupported: vi.fn(() => false) } },
      });
      expect(inspectWordHost()).toEqual({
        kind: "blocked",
        title: "Word needs annotation support",
        detail: "Verseform requires WordApi 1.7 and a connected Microsoft 365 subscription.",
      });

      vi.stubGlobal("Office", {
        HostType: { Word: "Word" },
        context: { host: "Word", requirements: { isSetSupported: vi.fn(() => true) } },
      });
      expect(inspectWordHost()).toMatchObject({ kind: "ready", title: "Word is ready" });
      expect(officeReadinessFailure()).toMatchObject({
        kind: "blocked",
        title: "Word could not start Verseform",
        detail: expect.stringMatching(/did not change your document/u),
      });
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

type OfficeRegistration = { remove: ReturnType<typeof vi.fn> };

function noopHandlers(): WordHostHandlers {
  return {
    onParagraphChanged: async () => undefined,
    onAnnotationActivated: async () => undefined,
    onAnnotationRemoved: async () => undefined,
  };
}

function eventSource(registration: OfficeRegistration): { add: ReturnType<typeof vi.fn> } {
  return { add: vi.fn(() => registration) };
}

const adapterAnnotation: AnnotatedReference = {
  annotationId: "annotation-1",
  paragraphId: "paragraph-1",
  paragraphRevision: 1,
  paragraphText: "Read John 3:16.",
  sourceText: "John 3:16",
  range: { from: 5, to: 14 },
  reference: { bookId: "JHN", bookName: "John", chapter: 3, verseStart: 16 },
  translationId: "VFW-010-FAKE",
};

function annotationCurrentFixture(overrides: {
  annotationId?: string;
  state?: string;
  paragraphText?: string;
  rangeText?: string;
  critiqueStart?: number;
  critiqueLength?: number;
  rangeParagraphId?: string;
} = {}): { run: ReturnType<typeof vi.fn>; wordAnnotation: { load: ReturnType<typeof vi.fn> } } {
  const range = {
    text: overrides.rangeText ?? adapterAnnotation.sourceText,
    paragraphs: {
      items: [{ uniqueLocalId: overrides.rangeParagraphId ?? adapterAnnotation.paragraphId }],
      load: vi.fn(),
    },
    load: vi.fn(),
  };
  const wordAnnotation = {
    id: overrides.annotationId ?? adapterAnnotation.annotationId,
    state: overrides.state ?? "Created",
    critiqueAnnotation: {
      range,
      critique: {
        start: overrides.critiqueStart ?? adapterAnnotation.range.from,
        length: overrides.critiqueLength ?? adapterAnnotation.range.to - adapterAnnotation.range.from,
      },
      load: vi.fn(),
    },
    load: vi.fn(),
  };
  const paragraph = {
    text: overrides.paragraphText ?? adapterAnnotation.paragraphText,
    uniqueLocalId: adapterAnnotation.paragraphId,
    load: vi.fn(),
  };
  const context = {
    document: {
      getParagraphByUniqueLocalId: vi.fn(() => paragraph),
      getAnnotationById: vi.fn(() => wordAnnotation),
    },
    sync: vi.fn(async () => undefined),
  };
  const run = vi.fn(async (...args: unknown[]) => {
    const batch = typeof args[0] === "function" ? args[0] : args[1];
    return (batch as (context: unknown) => Promise<unknown>)(context);
  });
  return { run, wordAnnotation };
}

describe("Office gateway adapter seam", () => {
  it("cleans partial registration, unregisters exact handlers through the retained document context, and attempts every removal", async () => {
    const first: OfficeRegistration = { remove: vi.fn() };
    const second: OfficeRegistration = { remove: vi.fn() };
    const partialDocument = {
      onParagraphChanged: eventSource(first),
      onAnnotationHovered: eventSource(second),
      onAnnotationClicked: { add: vi.fn(() => { throw new Error("third registration failed"); }) },
      onAnnotationRemoved: eventSource({ remove: vi.fn() }),
    };
    const partialContext = { document: partialDocument, sync: vi.fn(async () => undefined) };
    const partialRun = vi.fn(async (...args: unknown[]) => {
      const batch = typeof args[0] === "function" ? args[0] : args[1];
      return (batch as (context: unknown) => Promise<unknown>)(partialContext);
    });
    vi.stubGlobal("Word", { run: partialRun });
    try {
      await expect(new OfficeWordGateway().start(noopHandlers())).rejects.toThrow("third registration failed");
      expect(first.remove).toHaveBeenCalledTimes(1);
      expect(second.remove).toHaveBeenCalledTimes(1);
      expect(partialContext.sync).toHaveBeenCalledTimes(1);
    } finally {
      vi.unstubAllGlobals();
    }

    const registrations: OfficeRegistration[] = [
      { remove: vi.fn() },
      { remove: vi.fn(() => { throw new Error("one removal failed"); }) },
      { remove: vi.fn() },
      { remove: vi.fn() },
    ];
    const retainedDocument = {
      onParagraphChanged: eventSource(registrations[0]!),
      onAnnotationHovered: eventSource(registrations[1]!),
      onAnnotationClicked: eventSource(registrations[2]!),
      onAnnotationRemoved: eventSource(registrations[3]!),
    };
    const retainedContext = { document: retainedDocument, sync: vi.fn(async () => undefined) };
    const retainedRun = vi.fn(async (...args: unknown[]) => {
      const batch = typeof args[0] === "function" ? args[0] : args[1];
      return (batch as (context: unknown) => Promise<unknown>)(retainedContext);
    });
    vi.stubGlobal("Word", { run: retainedRun });
    try {
      const runtime = await new OfficeWordGateway().start(noopHandlers());
      await expect(runtime.stop()).rejects.toThrow("one removal failed");
      for (const registration of registrations) expect(registration.remove).toHaveBeenCalledTimes(1);
      expect(retainedRun.mock.calls[1]?.[0]).toBe(retainedDocument);
      expect(retainedContext.sync).toHaveBeenCalledTimes(2);
      await expect(runtime.stop()).rejects.toThrow("one removal failed");
      for (const registration of registrations) expect(registration.remove).toHaveBeenCalledTimes(1);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("loads annotation state and refuses a non-Created annotation before Range.insertText", async () => {
    const annotation: AnnotatedReference = {
      annotationId: "annotation-1",
      paragraphId: "paragraph-1",
      paragraphRevision: 1,
      paragraphText: "Read John 3:16.",
      sourceText: "John 3:16",
      range: { from: 5, to: 14 },
      reference: { bookId: "JHN", bookName: "John", chapter: 3, verseStart: 16 },
      translationId: "VFW-010-FAKE",
    };
    const range = {
      text: "John 3:16",
      paragraphs: { items: [{ uniqueLocalId: "paragraph-1" }], load: vi.fn() },
      load: vi.fn(),
      insertText: vi.fn(),
    };
    const wordAnnotation = {
      id: "annotation-1",
      state: "Accepted",
      critiqueAnnotation: { range, critique: { start: 5, length: 9 }, load: vi.fn() },
      load: vi.fn(),
    };
    const paragraph = { text: "Read John 3:16.", uniqueLocalId: "paragraph-1", load: vi.fn() };
    const context = {
      document: {
        getParagraphByUniqueLocalId: vi.fn(() => paragraph),
        getAnnotationById: vi.fn(() => wordAnnotation),
      },
      sync: vi.fn(async () => undefined),
    };
    const run = vi.fn(async (...args: unknown[]) => {
      const batch = typeof args[0] === "function" ? args[0] : args[1];
      return (batch as (context: unknown) => Promise<unknown>)(context);
    });
    vi.stubGlobal("Word", { run });
    try {
      await expect(new OfficeWordGateway().replace(annotation, "fake")).resolves.toBe("stale");
      expect(wordAnnotation.load).toHaveBeenCalledWith("id,state");
      expect(range.insertText).not.toHaveBeenCalled();
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("revalidates a retired ID through the Office adapter only when every native identity field is current", async () => {
    const valid = annotationCurrentFixture();
    vi.stubGlobal("Word", { run: valid.run });
    try {
      await expect(new OfficeWordGateway().isAnnotationCurrent(adapterAnnotation)).resolves.toBe(true);
      expect(valid.wordAnnotation.load).toHaveBeenCalledWith("id,state");

      for (const mismatch of [
        { annotationId: "different-id" },
        { state: "Accepted" },
        { rangeText: "John 3:15" },
        { paragraphText: "Read John 3:15." },
        { critiqueStart: 4 },
        { critiqueLength: 8 },
        { rangeParagraphId: "another-paragraph" },
      ]) {
        const invalid = annotationCurrentFixture(mismatch);
        vi.stubGlobal("Word", { run: invalid.run });
        await expect(new OfficeWordGateway().isAnnotationCurrent(adapterAnnotation)).resolves.toBe(false);
      }

      vi.stubGlobal("Word", { run: vi.fn(async () => { throw { code: "ItemNotFound" }; }) });
      await expect(new OfficeWordGateway().isAnnotationCurrent(adapterAnnotation)).resolves.toBe(false);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

class FakeElement {
  public readonly dataset: Record<string, string | undefined> = {};
  public textContent = "";
  public hidden = false;
  public disabled = false;
  public value = "";
  public children: FakeElement[] = [];
  public focusCalls = 0;
  private readonly listeners = new Map<string, () => void>();

  public addEventListener(type: string, listener: () => void): void {
    this.listeners.set(type, listener);
  }

  public focus(): void {
    this.focusCalls += 1;
  }

  public click(): void {
    this.listeners.get("click")?.();
  }

  public change(): void {
    this.listeners.get("change")?.();
  }

  public replaceChildren(...children: FakeElement[]): void {
    this.children = children;
  }
}

describe("task-pane DOM harness", () => {
  it("announces preview markup, catches UI action rejection, focuses status, and exposes Cancel", async () => {
    const source = readFileSync("index.html", "utf8");
    const styles = readFileSync("src/styles.css", "utf8");
    expect(source).toContain('name="color-scheme" content="light dark"');
    expect(source).toMatch(/class="status-card"[^>]*aria-live="polite"/u);
    expect(source).toMatch(/id="preview-card"[^>]*aria-live="polite"/u);
    expect(source).toContain('id="preview-attribution"');
    expect(source).toContain('id="translation-select"');
    expect(source).toContain('id="clear-cache-button"');
    expect(source).toContain('id="cancel-button"');
    expect(styles).toMatch(/color-scheme:\s*light dark/u);
    expect(styles).toContain("@media (prefers-color-scheme: dark)");
    expect(styles).toContain("@media (forced-colors: active)");
    expect(styles).toMatch(/--canvas:\s*Canvas/u);
    expect(styles).toMatch(/--text:\s*CanvasText/u);

    const elements = new Map<string, FakeElement>([
      [".status-card", new FakeElement()], ["#status-title", new FakeElement()], ["#status-detail", new FakeElement()],
      ["#preview-card", new FakeElement()], ["#preview-heading", new FakeElement()], ["#preview-text", new FakeElement()],
      ["#preview-attribution", new FakeElement()], ["#insert-button", new FakeElement()], ["#cancel-button", new FakeElement()],
      ["#translation-select", new FakeElement()], ["#clear-cache-button", new FakeElement()],
    ]);
    vi.stubGlobal("document", {
      querySelector: (selector: string) => elements.get(selector) ?? null,
      createElement: () => new FakeElement(),
    });
    try {
      const pane = mountTaskPane();
      let cancellations = 0;
      let selectedTranslation = "";
      let cacheClears = 0;
      pane.onCancel(async () => { cancellations += 1; });
      pane.onTranslationChange(async (translationId) => { selectedTranslation = translationId; });
      pane.onClearCache(async () => { cacheClears += 1; });
      pane.render({
        phase: "preview", title: "Preview", detail: "Fake", canInsert: true, canCancel: true,
        catalogPhase: "ready", translations: dbsTranslations, selectedTranslationId: "ENGNASB",
        canSelectTranslation: true, canClearCache: true,
        preview: {
          heading: "John 3:16", text: "Fake", attribution: "Local test data.", insertText: "Fake",
          translationId: "VFW-010-FAKE", translationName: "VFW-010 test text",
          citationLabel: "VFW-010 test text", cached: false,
        },
      });
      elements.get("#cancel-button")?.click();
      await Promise.resolve();
      expect(cancellations).toBe(1);
      expect(elements.get("#preview-card")?.hidden).toBe(false);
      expect(elements.get("#preview-attribution")?.textContent).toBe("Local test data.");
      expect(elements.get("#insert-button")?.disabled).toBe(false);
      expect(elements.get("#translation-select")?.children).toHaveLength(2);
      expect(elements.get("#translation-select")?.value).toBe("ENGNASB");
      expect(elements.get("#translation-select")?.disabled).toBe(false);
      elements.get("#translation-select")!.value = "ENGTEST";
      elements.get("#translation-select")?.change();
      elements.get("#clear-cache-button")?.click();
      await Promise.resolve();
      expect(selectedTranslation).toBe("ENGTEST");
      expect(cacheClears).toBe(1);

      pane.render({ phase: "watching", title: "Refused", detail: "No change", canInsert: false, canCancel: false, focusTarget: "status" });
      expect(elements.get("#status-title")?.focusCalls).toBe(1);

      pane.renderHostStatus("blocked", "Word needs annotation support", "WordApi 1.7 is required.");
      expect(elements.get(".status-card")?.dataset.kind).toBe("blocked");
      expect(elements.get("#insert-button")?.disabled).toBe(true);
      expect(elements.get("#cancel-button")?.disabled).toBe(true);
      expect(elements.get("#status-title")?.focusCalls).toBe(2);

      pane.onInsert(async () => { throw new Error("UI action failure"); });
      pane.render({
        phase: "preview", title: "Preview", detail: "Fake", canInsert: true, canCancel: true,
        preview: {
          heading: "John", text: "Fake", attribution: "Local test data.", insertText: "Fake",
          translationId: "VFW-010-FAKE", translationName: "VFW-010 test text",
          citationLabel: "VFW-010 test text", cached: false,
        },
      });
      elements.get("#insert-button")?.click();
      await Promise.resolve();
      expect(elements.get(".status-card")?.dataset.kind).toBe("blocked");
      expect(elements.get("#status-detail")?.textContent).toMatch(/could not confirm the result\. Inspect the document/u);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
