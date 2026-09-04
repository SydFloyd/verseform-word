import { describe, expect, it, vi } from "vitest";
import {
  DBS_ORIGIN,
  DbsScriptureProvider,
  FetchDbsTransport,
  type DbsTransport,
} from "../src/adapters/dbsScriptureProvider";
import {
  BrowserScriptureCache,
  CHAPTER_CACHE_TTL_MS,
  SCRIPTURE_CACHE_MAX_BYTES,
  SCRIPTURE_CACHE_MAX_ENTRIES,
  SCRIPTURE_CACHE_STORAGE_KEY,
} from "../src/adapters/scriptureCache";
import {
  BrowserTranslationPreference,
  TRANSLATION_PREFERENCE_KEY,
} from "../src/adapters/translationPreference";
import {
  DBS_CHAPTER_LIMIT_BYTES,
  DBS_SECTION_BOOK_CODES,
  normalizeDbsVerseText,
  parseDbsCatalog,
  parseDbsChapter,
  previewForPassage,
  selectInitialTranslation,
} from "../src/core/scripture";
import { STANDARD_CANON } from "../src/core/canon";

const catalogBody = JSON.stringify([
  {
    abbr: "ENGTEST",
    title: "DBS Test Bible",
    title_vernacular: "Test Bible",
    iso: "eng",
    script: "Latn",
    year: 2026,
    copyright: "DBS test fixture — not production Scripture.",
  },
  {
    abbr: "ENGNASB",
    title: "New American Standard Bible",
    iso: "eng",
    script: "Latn",
    year: 1960,
    copyright: "© The Lockman Foundation",
  },
]);

const john3Body = JSON.stringify([{
  "JN3.16": "“For  God so loved the \n world, that He gave His only begotten Son.",
  "JN3.17": "For God did not send the Son into the world to judge the world.",
}]);

const john316 = {
  bookId: "JHN",
  bookName: "John",
  chapter: 3,
  verseStart: 16,
} as const;

class MemoryStorage {
  public readonly values = new Map<string, string>();
  public getItem(key: string): string | null { return this.values.get(key) ?? null; }
  public setItem(key: string, value: string): void { this.values.set(key, value); }
  public removeItem(key: string): void { this.values.delete(key); }
}

class RecordedTransport implements DbsTransport {
  public catalogCalls = 0;
  public chapterCalls: Array<[string, string, number]> = [];
  public async getCatalog(): Promise<string> {
    this.catalogCalls += 1;
    return catalogBody;
  }
  public async getChapter(translationId: string, bookId: string, chapter: number): Promise<string> {
    this.chapterCalls.push([translationId, bookId, chapter]);
    return john3Body;
  }
}

describe("DBS pure contracts", () => {
  it("parses the recorded catalog, derives citation labels, and prefers saved/NASB deterministically", () => {
    const translations = parseDbsCatalog(catalogBody);
    expect(translations).toHaveLength(2);
    expect(translations.find((item) => item.id === "ENGNASB")).toMatchObject({
      citationLabel: "NASB",
      name: "New American Standard Bible",
      year: "1960",
      attribution: "New American Standard Bible (NASB): © The Lockman Foundation",
    });
    expect(selectInitialTranslation(translations)?.id).toBe("ENGNASB");
    expect(selectInitialTranslation(translations, "ENGTEST")?.id).toBe("ENGTEST");
    expect(() => parseDbsCatalog('[{"abbr":"../NASB","title":"Unsafe"}]')).toThrow(/invalid/u);
  });

  it("pins one unambiguous DBS section code for every standard-canon book", () => {
    expect(Object.keys(DBS_SECTION_BOOK_CODES)).toEqual(
      STANDARD_CANON.books.map((book) => book.id),
    );
    expect(new Set(Object.values(DBS_SECTION_BOOK_CODES))).toHaveLength(66);
  });

  it("repairs DBS serialization boundaries and removes a trailing section heading", () => {
    expect(normalizeDbsVerseText(
      "James, a  bond-servant of \n God and of the Lord Jesus Christ,To the twelve tribes.",
    )).toBe("James, a bond-servant of God and of the Lord Jesus Christ, To the twelve tribes.");
    expect(normalizeDbsVerseText(
      "Therefore, to one who knows the right thing to do, to him it is sin.Misuse of Riches",
    )).toBe("Therefore, to one who knows the right thing to do, to him it is sin.");

    const chapter = parseDbsChapter(john3Body, 3, "JHN");
    expect(chapter.get(16)).toBe("“For God so loved the world, that He gave His only begotten Son.");
    expect(parseDbsChapter('[{"S13.1":"Numbered-book text."}]', 3, "1SA").get(1))
      .toBe("Numbered-book text.");
    expect(parseDbsChapter("[{}]", 3, "JHN")).toEqual(new Map());
    expect(() => parseDbsChapter('[{"MK3.16":"wrong book"}]', 3, "JHN"))
      .toThrow(/different coordinates/u);
    expect(() => parseDbsChapter('[{"JN4.16":"wrong chapter"}]', 3, "JHN"))
      .toThrow(/different coordinates/u);
    expect(() => parseDbsChapter("[{},{}]", 3, "JHN")).toThrow(/invalid verse entry/u);
  });

  it("builds editable passage, citation, and provider attribution text", () => {
    const preview = previewForPassage({
      reference: john316,
      display: "John 3:16",
      translationId: "ENGNASB",
      citationLabel: "NASB",
      translationName: "New American Standard Bible",
      attribution: "New American Standard Bible (NASB): © The Lockman Foundation",
      text: "For God so loved the world.",
      cached: false,
    });
    expect(preview.insertText).toBe(
      "For God so loved the world. (John 3:16, NASB)\nNew American Standard Bible (NASB): © The Lockman Foundation",
    );
  });
});

describe("DBS adapter boundary", () => {
  it("caches only validated catalog/chapter bodies and selects a range locally", async () => {
    const transport = new RecordedTransport();
    const storage = new MemoryStorage();
    const cache = new BrowserScriptureCache(storage, () => 1_000);
    const provider = new DbsScriptureProvider(transport, cache);

    await expect(provider.listTranslations()).resolves.toMatchObject({ cached: false });
    await expect(provider.getPassage(
      { ...john316, verseEnd: 17 },
      "ENGNASB",
    )).resolves.toMatchObject({
      display: "John 3:16-17",
      translationId: "ENGNASB",
      text: expect.stringContaining("For God did not send"),
      cached: false,
    });
    expect(transport.chapterCalls).toEqual([["ENGNASB", "JHN", 3]]);

    const second = new DbsScriptureProvider(transport, cache);
    await expect(second.listTranslations()).resolves.toMatchObject({ cached: true });
    await expect(second.getPassage(john316, "ENGNASB")).resolves.toMatchObject({ cached: true });
    expect(transport.catalogCalls).toBe(1);
    expect(transport.chapterCalls).toHaveLength(1);
    expect(JSON.stringify(storage.values)).not.toContain("surrounding document prose");
  });

  it("does not persist a live response until its schema is valid", async () => {
    const storage = new MemoryStorage();
    const cache = new BrowserScriptureCache(storage, () => 1_000);
    const provider = new DbsScriptureProvider({
      async getCatalog() { return '[{"abbr":"../unsafe","title":"Changed"}]'; },
      async getChapter() { return "[]"; },
    }, cache);
    await expect(provider.listTranslations()).rejects.toThrow(/invalid/u);
    expect(storage.getItem(SCRIPTURE_CACHE_STORAGE_KEY)).toBeNull();
  });

  it("constructs only allowlisted coordinate URLs with no request body or credentials", async () => {
    const request = vi.fn(async (_input: RequestInfo | URL, _init?: RequestInit) => new Response(john3Body, {
      status: 200,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    }));
    const transport = new FetchDbsTransport(request, 500);
    await expect(transport.getChapter("ENGNASB", "JHN", 3)).resolves.toBe(john3Body);
    expect(request).toHaveBeenCalledTimes(1);
    const [url, init] = request.mock.calls[0]!;
    expect(String(url)).toBe(`${DBS_ORIGIN}/api/bible-text/ENGNASB/JHN/3`);
    expect(init).toMatchObject({
      method: "GET",
      credentials: "omit",
      redirect: "error",
      referrerPolicy: "no-referrer",
      cache: "no-store",
    });
    expect(init?.body).toBeUndefined();
    expect(JSON.stringify(init)).not.toContain("document prose");

    await expect(transport.getChapter("../NASB", "JHN", 3)).rejects.toThrow(/invalid/u);
    expect(request).toHaveBeenCalledTimes(1);
  });

  it("rejects oversized or non-JSON responses before parsing", async () => {
    const oversized = new FetchDbsTransport(vi.fn(async () => new Response("[]", {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Content-Length": String(DBS_CHAPTER_LIMIT_BYTES + 1),
      },
    })));
    await expect(oversized.getChapter("ENGNASB", "JHN", 3)).rejects.toThrow(/safety limit/u);

    const html = new FetchDbsTransport(vi.fn(async () => new Response("<html></html>", {
      status: 200,
      headers: { "Content-Type": "text/html" },
    })));
    await expect(html.getCatalog()).rejects.toThrow(/unexpected content type/u);
  });

  it("propagates user cancellation through the browser request signal", async () => {
    let transportSignal: AbortSignal | undefined;
    const request = vi.fn(async (_input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
      transportSignal = init?.signal ?? undefined;
      return new Promise<Response>((_resolve, reject) => {
        transportSignal?.addEventListener(
          "abort",
          () => reject(new DOMException("cancelled", "AbortError")),
          { once: true },
        );
      });
    });
    const controller = new AbortController();
    const pending = new FetchDbsTransport(request).getCatalog(controller.signal);
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: "AbortError" });
    expect(transportSignal?.aborted).toBe(true);
  });
});

describe("bounded local Scripture storage", () => {
  it("expires chapters, enforces entry/byte limits, clears content, and keeps preference separate", async () => {
    const storage = new MemoryStorage();
    let now = 10_000;
    const cache = new BrowserScriptureCache(storage, () => now);
    await cache.set({ kind: "catalog", key: "catalog" }, catalogBody);
    await cache.set({ kind: "chapter", key: "ENGNASB/JHN/3" }, john3Body);
    expect(await cache.get({ kind: "chapter", key: "ENGNASB/JHN/3" })).toBe(john3Body);

    for (let index = 1; index <= SCRIPTURE_CACHE_MAX_ENTRIES + 8; index += 1) {
      now += 1;
      await cache.set(
        { kind: "chapter", key: `ENGNASB/JHN/${index}` },
        JSON.stringify([{ [`JN${index}.1`]: "x".repeat(70_000) }]),
      );
    }
    const raw = storage.getItem(SCRIPTURE_CACHE_STORAGE_KEY) ?? "";
    const root = JSON.parse(raw) as { version: number; entries: unknown[] };
    expect(root.version).toBe(1);
    expect(root.entries.length).toBeLessThanOrEqual(SCRIPTURE_CACHE_MAX_ENTRIES);
    expect(new TextEncoder().encode(raw).byteLength).toBeLessThanOrEqual(SCRIPTURE_CACHE_MAX_BYTES);

    now += CHAPTER_CACHE_TTL_MS + 1;
    expect(await cache.get({ kind: "chapter", key: "ENGNASB/JHN/3" })).toBeUndefined();

    const preference = new BrowserTranslationPreference(storage);
    await preference.save("ENGNASB");
    await cache.clear();
    expect(storage.getItem(SCRIPTURE_CACHE_STORAGE_KEY)).toBeNull();
    expect(storage.getItem(TRANSLATION_PREFERENCE_KEY)).toBe("ENGNASB");
    await expect(preference.load()).resolves.toBe("ENGNASB");
  });

  it("discards malformed cache state without blocking future live data", async () => {
    const storage = new MemoryStorage();
    storage.setItem(SCRIPTURE_CACHE_STORAGE_KEY, '{"version":99,"entries":[]}');
    const cache = new BrowserScriptureCache(storage, () => 1_000);
    await expect(cache.get({ kind: "catalog", key: "catalog" })).resolves.toBeUndefined();
    expect(storage.getItem(SCRIPTURE_CACHE_STORAGE_KEY)).toBeNull();
    await cache.set({ kind: "catalog", key: "catalog" }, catalogBody);
    await expect(cache.get({ kind: "catalog", key: "catalog" })).resolves.toBe(catalogBody);
  });

  it("discards future-dated cache activity instead of preserving it during pruning", async () => {
    const storage = new MemoryStorage();
    storage.setItem(SCRIPTURE_CACHE_STORAGE_KEY, JSON.stringify({
      version: 1,
      entries: [{
        kind: "chapter",
        key: "ENGNASB/JHN/3",
        fetchedAtMs: 900,
        accessedAtMs: 1_100,
        body: john3Body,
      }],
    }));
    const cache = new BrowserScriptureCache(storage, () => 1_000);
    await expect(cache.get({ kind: "chapter", key: "ENGNASB/JHN/3" })).resolves.toBeUndefined();
    expect(JSON.parse(storage.getItem(SCRIPTURE_CACHE_STORAGE_KEY) ?? "{}").entries).toEqual([]);
  });
});
