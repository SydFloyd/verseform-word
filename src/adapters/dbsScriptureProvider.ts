import type { NormalizedReference } from "../core/reference";
import {
  DBS_CATALOG_LIMIT_BYTES,
  DBS_CHAPTER_LIMIT_BYTES,
  isTranslationId,
  parseDbsCatalog,
  parseDbsChapter,
  passageDisplay,
  type Passage,
  type ScriptureProvider,
  type Translation,
  type TranslationCatalog,
} from "../core/scripture";
import type { ScriptureCache } from "./scriptureCache";

export const DBS_ORIGIN = "https://arc.dbs.org";
const REQUEST_TIMEOUT_MS = 8_000;
const bookIdPattern = /^[1-3]?[A-Z]{1,3}$/u;

export interface DbsTransport {
  getCatalog(signal?: AbortSignal): Promise<string>;
  getChapter(
    translationId: string,
    bookId: string,
    chapter: number,
    signal?: AbortSignal,
  ): Promise<string>;
}

async function boundedResponseBody(response: Response, limit: number): Promise<string> {
  if (!response.ok) throw new Error(`DBS returned HTTP ${response.status}.`);
  const contentType = response.headers.get("content-type")?.toLowerCase() ?? "";
  if (!contentType.includes("application/json")) {
    throw new Error("DBS returned an unexpected content type.");
  }
  const declaredLength = Number(response.headers.get("content-length"));
  if (Number.isFinite(declaredLength) && declaredLength > limit) {
    throw new Error("The DBS response exceeded Verseform's safety limit.");
  }
  if (!response.body) {
    const body = await response.text();
    if (new TextEncoder().encode(body).byteLength > limit) {
      throw new Error("The DBS response exceeded Verseform's safety limit.");
    }
    return body;
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  let received = 0;
  let body = "";
  try {
    while (true) {
      const result = await reader.read();
      if (result.done) break;
      received += result.value.byteLength;
      if (received > limit) {
        await reader.cancel();
        throw new Error("The DBS response exceeded Verseform's safety limit.");
      }
      try {
        body += decoder.decode(result.value, { stream: true });
      } catch {
        throw new Error("DBS returned text in an unsupported encoding.");
      }
    }
    try {
      body += decoder.decode();
    } catch {
      throw new Error("DBS returned text in an unsupported encoding.");
    }
    return body;
  } finally {
    reader.releaseLock();
  }
}

export class FetchDbsTransport implements DbsTransport {
  public constructor(
    private readonly request: typeof fetch = globalThis.fetch.bind(globalThis),
    private readonly timeoutMs = REQUEST_TIMEOUT_MS,
  ) {}

  public async getCatalog(signal?: AbortSignal): Promise<string> {
    return this.get(new URL("/api/bible-text/", DBS_ORIGIN), DBS_CATALOG_LIMIT_BYTES, signal);
  }

  public async getChapter(
    translationId: string,
    bookId: string,
    chapter: number,
    signal?: AbortSignal,
  ): Promise<string> {
    if (!isTranslationId(translationId)
      || !bookIdPattern.test(bookId)
      || !Number.isSafeInteger(chapter)
      || chapter < 1
      || chapter > 200) throw new Error("The Scripture request was invalid.");
    const path = `/api/bible-text/${encodeURIComponent(translationId)}/${encodeURIComponent(bookId)}/${chapter}`;
    return this.get(new URL(path, DBS_ORIGIN), DBS_CHAPTER_LIMIT_BYTES, signal);
  }

  private async get(url: URL, limit: number, signal?: AbortSignal): Promise<string> {
    const controller = new AbortController();
    let timedOut = false;
    const timeout = globalThis.setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, this.timeoutMs);
    const cancel = () => controller.abort();
    if (signal?.aborted) cancel();
    else signal?.addEventListener("abort", cancel, { once: true });
    try {
      const response = await this.request(url, {
        method: "GET",
        headers: { Accept: "application/json" },
        credentials: "omit",
        redirect: "error",
        referrerPolicy: "no-referrer",
        cache: "no-store",
        signal: controller.signal,
      });
      return await boundedResponseBody(response, limit);
    } catch (error) {
      if (signal?.aborted) throw new DOMException("DBS request cancelled.", "AbortError");
      if (timedOut) throw new Error("DBS did not respond in time. Check your connection and try again.");
      if (error instanceof Error
        && (error.message.startsWith("DBS") || error.message.startsWith("The DBS"))) throw error;
      throw new Error("DBS is unavailable. Check your connection and try again.");
    } finally {
      globalThis.clearTimeout(timeout);
      signal?.removeEventListener("abort", cancel);
    }
  }
}

export class DbsScriptureProvider implements ScriptureProvider {
  private catalog?: Translation[];

  public constructor(
    private readonly transport: DbsTransport,
    private readonly cache: ScriptureCache,
  ) {}

  public async listTranslations(signal?: AbortSignal): Promise<TranslationCatalog> {
    const cachedBody = await this.cache.get({ kind: "catalog", key: "catalog" });
    if (cachedBody !== undefined) {
      try {
        this.catalog = parseDbsCatalog(cachedBody);
        return { translations: this.catalog, cached: true };
      } catch {
        await this.cache.clear().catch(() => undefined);
      }
    }
    const body = await this.transport.getCatalog(signal);
    signal?.throwIfAborted();
    this.catalog = parseDbsCatalog(body);
    signal?.throwIfAborted();
    await this.cache.set({ kind: "catalog", key: "catalog" }, body);
    return { translations: this.catalog, cached: false };
  }

  public async getPassage(
    reference: NormalizedReference,
    translationId: string,
    signal?: AbortSignal,
  ): Promise<Passage> {
    if (!isTranslationId(translationId)) throw new Error("That translation identifier is invalid.");
    const translation = this.catalog?.find((item) => item.id === translationId);
    if (!translation) throw new Error("That translation is not in the authorized DBS catalog.");
    const cacheKey = `${translationId}/${reference.bookId}/${reference.chapter}`;
    let body = await this.cache.get({ kind: "chapter", key: cacheKey });
    let cached = body !== undefined;
    if (body !== undefined) {
      try {
        parseDbsChapter(body, reference.chapter, reference.bookId);
      } catch {
        await this.cache.clear().catch(() => undefined);
        body = undefined;
        cached = false;
      }
    }
    if (body === undefined) {
      body = await this.transport.getChapter(
        translationId,
        reference.bookId,
        reference.chapter,
        signal,
      );
      signal?.throwIfAborted();
      parseDbsChapter(body, reference.chapter, reference.bookId);
      signal?.throwIfAborted();
      await this.cache.set({ kind: "chapter", key: cacheKey }, body);
    }
    const chapter = parseDbsChapter(body, reference.chapter, reference.bookId);
    const end = reference.verseEnd ?? reference.verseStart;
    const selected: string[] = [];
    for (let verse = reference.verseStart; verse <= end; verse += 1) {
      const text = chapter.get(verse);
      if (!text) {
        throw new Error(`${passageDisplay(reference)} is unavailable in ${translation.name}.`);
      }
      selected.push(text);
    }
    return {
      reference,
      display: passageDisplay(reference),
      translationId: translation.id,
      citationLabel: translation.citationLabel,
      translationName: translation.name,
      attribution: translation.attribution,
      text: selected.join(" "),
      cached,
    };
  }

  public async clearCache(): Promise<void> {
    await this.cache.clear();
  }
}
