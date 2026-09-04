export const SCRIPTURE_CACHE_STORAGE_KEY = "verseform-word.scripture-cache.v1";
export const SCRIPTURE_CACHE_SCHEMA_VERSION = 1;
export const SCRIPTURE_CACHE_MAX_ENTRIES = 65;
export const SCRIPTURE_CACHE_MAX_BYTES = 4 * 1024 * 1024;
export const CATALOG_CACHE_TTL_MS = 24 * 60 * 60 * 1_000;
export const CHAPTER_CACHE_TTL_MS = 7 * 24 * 60 * 60 * 1_000;

type CacheKind = "catalog" | "chapter";
type CacheEntry = {
  key: string;
  kind: CacheKind;
  fetchedAtMs: number;
  accessedAtMs: number;
  body: string;
};
type CacheRoot = { version: 1; entries: CacheEntry[] };

export type ScriptureCacheKey = { kind: CacheKind; key: string };

export interface ScriptureCache {
  get(input: ScriptureCacheKey): Promise<string | undefined>;
  set(input: ScriptureCacheKey, body: string): Promise<void>;
  clear(): Promise<void>;
}

type StorageLike = Pick<Storage, "getItem" | "setItem" | "removeItem">;

function bytes(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function entryLimit(kind: CacheKind): number {
  return kind === "catalog" ? 8 * 1024 * 1024 : 2 * 1024 * 1024;
}

function entryTtl(kind: CacheKind): number {
  return kind === "catalog" ? CATALOG_CACHE_TTL_MS : CHAPTER_CACHE_TTL_MS;
}

function validKey(input: ScriptureCacheKey): boolean {
  return input.kind === "catalog"
    ? input.key === "catalog"
    : /^[A-Za-z0-9_-]{1,64}\/[1-3]?[A-Z]{1,3}\/(?:[1-9]\d{0,2})$/u.test(input.key);
}

function parseRoot(raw: string | null): CacheRoot | undefined {
  if (!raw || bytes(raw) > SCRIPTURE_CACHE_MAX_BYTES) return undefined;
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return undefined;
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const root = value as Partial<CacheRoot>;
  if (root.version !== SCRIPTURE_CACHE_SCHEMA_VERSION
    || !Array.isArray(root.entries)
    || root.entries.length > SCRIPTURE_CACHE_MAX_ENTRIES) return undefined;
  const seen = new Set<string>();
  const entries: CacheEntry[] = [];
  for (const candidate of root.entries) {
    if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) return undefined;
    const entry = candidate as Partial<CacheEntry>;
    if ((entry.kind !== "catalog" && entry.kind !== "chapter")
      || typeof entry.key !== "string"
      || !validKey({ kind: entry.kind, key: entry.key })
      || seen.has(`${entry.kind}:${entry.key}`)
      || typeof entry.body !== "string"
      || bytes(entry.body) > entryLimit(entry.kind)
      || typeof entry.fetchedAtMs !== "number"
      || !Number.isSafeInteger(entry.fetchedAtMs)
      || entry.fetchedAtMs < 0
      || typeof entry.accessedAtMs !== "number"
      || !Number.isSafeInteger(entry.accessedAtMs)
      || entry.accessedAtMs < entry.fetchedAtMs) return undefined;
    seen.add(`${entry.kind}:${entry.key}`);
    entries.push(entry as CacheEntry);
  }
  return { version: 1, entries };
}

export class BrowserScriptureCache implements ScriptureCache {
  public constructor(
    private readonly storage: StorageLike,
    private readonly now: () => number = () => Date.now(),
  ) {}

  public async get(input: ScriptureCacheKey): Promise<string | undefined> {
    if (!validKey(input)) return undefined;
    const root = this.read();
    const currentTime = this.now();
    const retained = root.entries.filter((entry) => (
      currentTime >= entry.fetchedAtMs
      && currentTime >= entry.accessedAtMs
      && currentTime - entry.fetchedAtMs <= entryTtl(entry.kind)
    ));
    const entry = retained.find((candidate) => (
      candidate.kind === input.kind && candidate.key === input.key
    ));
    if (entry) entry.accessedAtMs = currentTime;
    if (retained.length !== root.entries.length || entry) this.write({ version: 1, entries: retained });
    return entry?.body;
  }

  public async set(input: ScriptureCacheKey, body: string): Promise<void> {
    if (!validKey(input) || bytes(body) > entryLimit(input.kind)) return;
    const currentTime = this.now();
    const root = this.read();
    const entries = root.entries.filter((entry) => (
      !(entry.kind === input.kind && entry.key === input.key)
      && currentTime >= entry.fetchedAtMs
      && currentTime >= entry.accessedAtMs
      && currentTime - entry.fetchedAtMs <= entryTtl(entry.kind)
    ));
    entries.push({ ...input, fetchedAtMs: currentTime, accessedAtMs: currentTime, body });
    entries.sort((left, right) => right.accessedAtMs - left.accessedAtMs);
    while (entries.length > SCRIPTURE_CACHE_MAX_ENTRIES
      || bytes(JSON.stringify({ version: 1, entries })) > SCRIPTURE_CACHE_MAX_BYTES) {
      entries.pop();
    }
    this.write({ version: 1, entries });
  }

  public async clear(): Promise<void> {
    try {
      this.storage.removeItem(SCRIPTURE_CACHE_STORAGE_KEY);
    } catch {
      throw new Error("The local Scripture cache could not be cleared.");
    }
  }

  private read(): CacheRoot {
    try {
      const raw = this.storage.getItem(SCRIPTURE_CACHE_STORAGE_KEY);
      const parsed = parseRoot(raw);
      if (parsed) return parsed;
      if (raw !== null) this.storage.removeItem(SCRIPTURE_CACHE_STORAGE_KEY);
    } catch {
      // A disabled or full browser store must not prevent a live DBS request.
    }
    return { version: 1, entries: [] };
  }

  private write(root: CacheRoot): void {
    try {
      this.storage.setItem(SCRIPTURE_CACHE_STORAGE_KEY, JSON.stringify(root));
    } catch {
      // Cache writes are an optimization. The validated live result remains usable.
    }
  }
}

export class MemoryScriptureCache implements ScriptureCache {
  private readonly values = new Map<string, string>();

  public async get(input: ScriptureCacheKey): Promise<string | undefined> {
    return this.values.get(`${input.kind}:${input.key}`);
  }
  public async set(input: ScriptureCacheKey, body: string): Promise<void> {
    this.values.set(`${input.kind}:${input.key}`, body);
  }
  public async clear(): Promise<void> {
    this.values.clear();
  }
}
