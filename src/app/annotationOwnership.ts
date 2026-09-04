/**
 * Persistent ownership contains only Verseform-generated annotation IDs and
 * small bounded bookkeeping. It never contains paragraph text, document
 * identity, account data, or Scripture content.
 */
export type AnnotationOwnershipEntry = {
  annotationId: string;
  missingSweeps: number;
};

export const MAX_OWNED_ANNOTATIONS = 256;
export const MAX_MISSING_SWEEPS = 8;
const RESERVATION_LIFETIME_MS = 60_000;

type Reservation = {
  reservationId: string;
  expiresAt: number;
};

type StoredOwnership = AnnotationOwnershipEntry | Reservation;

export type AnnotationRemovalUpdate = {
  removedAnnotationIds: readonly string[];
  missingAnnotationIds: readonly string[];
};

/**
 * Every mutation is a small storage transaction. It never spans a Word call:
 * callers reserve capacity, call Word, then commit or cancel.
 */
export interface AnnotationOwnership {
  load(): Promise<readonly AnnotationOwnershipEntry[]>;
  reserve(): Promise<string | undefined>;
  commit(reservationId: string, annotationId: string): Promise<boolean>;
  cancel(reservationId: string): Promise<void>;
  release(annotationId: string): Promise<void>;
  recordRemoval(update: AnnotationRemovalUpdate): Promise<void>;
}

function isAnnotation(entry: StoredOwnership): entry is AnnotationOwnershipEntry {
  return "annotationId" in entry;
}

function validAnnotation(entry: AnnotationOwnershipEntry): boolean {
  return typeof entry.annotationId === "string"
    && entry.annotationId.length > 0
    && entry.annotationId.length <= 200
    && Number.isInteger(entry.missingSweeps)
    && entry.missingSweeps >= 0
    && entry.missingSweeps <= MAX_MISSING_SWEEPS;
}

function validReservation(entry: Reservation): boolean {
  return typeof entry.reservationId === "string"
    && entry.reservationId.length > 0
    && entry.reservationId.length <= 200
    && Number.isFinite(entry.expiresAt);
}

/**
 * The legacy VFW-010 array of strings is intentionally migrated. Any other
 * malformed, duplicate, or oversized browser value is rejected without a
 * write so an existing recoverable host annotation is never forgotten.
 */
function strictStoredOwnership(raw: unknown): StoredOwnership[] {
  if (!Array.isArray(raw) || raw.length > MAX_OWNED_ANNOTATIONS) {
    throw new Error("Verseform annotation ownership is malformed");
  }
  if (raw.every((entry) => typeof entry === "string")) {
    const annotations = raw.map((annotationId) => ({ annotationId, missingSweeps: 0 }));
    if (!annotations.every(validAnnotation)
      || new Set(annotations.map((entry) => entry.annotationId)).size !== annotations.length) {
      throw new Error("Verseform annotation ownership has invalid or duplicate IDs");
    }
    return annotations;
  }

  const records: StoredOwnership[] = raw.map((entry): StoredOwnership => {
    if (typeof entry !== "object" || entry === null) {
      throw new Error("Verseform annotation ownership has an invalid entry");
    }
    const candidate = entry as Partial<AnnotationOwnershipEntry & Reservation>;
    // Migrate the first structured format without a discriminator.
    if ("annotationId" in candidate && !("reservationId" in candidate)) {
      const annotation = {
        annotationId: candidate.annotationId,
        missingSweeps: candidate.missingSweeps,
      } as AnnotationOwnershipEntry;
      if (!validAnnotation(annotation)) throw new Error("Verseform annotation ownership has an invalid entry");
      return annotation;
    }
    if ("reservationId" in candidate && !("annotationId" in candidate)) {
      const reservation = {
        reservationId: candidate.reservationId,
        expiresAt: candidate.expiresAt,
      } as Reservation;
      if (!validReservation(reservation)) throw new Error("Verseform annotation ownership has an invalid reservation");
      return reservation;
    }
    throw new Error("Verseform annotation ownership has a mixed entry");
  });

  const keys = records.map((record) => isAnnotation(record)
    ? `annotation:${record.annotationId}`
    : `reservation:${record.reservationId}`);
  if (new Set(keys).size !== keys.length) {
    throw new Error("Verseform annotation ownership has duplicate entries");
  }
  return records;
}

function withoutExpiredReservations(records: readonly StoredOwnership[], now: number): StoredOwnership[] {
  return records.filter((record) => isAnnotation(record) || record.expiresAt > now);
}

function entries(records: readonly StoredOwnership[]): AnnotationOwnershipEntry[] {
  return records.filter(isAnnotation).map((entry) => ({ ...entry }));
}

function reservationId(): string {
  const randomId = globalThis.crypto?.randomUUID;
  if (typeof randomId !== "function") throw new Error("Secure reservation IDs are unavailable");
  return `reservation:${randomId.call(globalThis.crypto)}`;
}

abstract class TransactionalAnnotationOwnership implements AnnotationOwnership {
  public abstract load(): Promise<readonly AnnotationOwnershipEntry[]>;
  public abstract reserve(): Promise<string | undefined>;
  public abstract commit(reservationId: string, annotationId: string): Promise<boolean>;
  public abstract cancel(reservationId: string): Promise<void>;
  public abstract release(annotationId: string): Promise<void>;
  public abstract recordRemoval(update: AnnotationRemovalUpdate): Promise<void>;

  protected nextReservation(records: readonly StoredOwnership[]): string | undefined {
    if (records.length >= MAX_OWNED_ANNOTATIONS) return undefined;
    return reservationId();
  }

  protected commitReservation(
    records: readonly StoredOwnership[],
    reservedId: string,
    annotationId: string,
  ): { records: StoredOwnership[]; committed: boolean } {
    if (!validAnnotation({ annotationId, missingSweeps: 0 })) return { records: [...records], committed: false };
    const reservationIndex = records.findIndex((record) => !isAnnotation(record) && record.reservationId === reservedId);
    if (reservationIndex < 0 || records.some((record) => isAnnotation(record) && record.annotationId === annotationId)) {
      return { records: [...records], committed: false };
    }
    const next = [...records];
    next.splice(reservationIndex, 1, { annotationId, missingSweeps: 0 });
    return { records: next, committed: true };
  }

  protected applyRemoval(records: readonly StoredOwnership[], update: AnnotationRemovalUpdate): StoredOwnership[] {
    const removed = new Set(update.removedAnnotationIds);
    const missing = new Set(update.missingAnnotationIds);
    return records.flatMap((record): StoredOwnership[] => {
      if (!isAnnotation(record)) return [record];
      if (removed.has(record.annotationId)) return [];
      if (!missing.has(record.annotationId)) return [record];
      if (record.missingSweeps + 1 >= MAX_MISSING_SWEEPS) return [];
      return [{ ...record, missingSweeps: record.missingSweeps + 1 }];
    });
  }
}

export class MemoryAnnotationOwnership extends TransactionalAnnotationOwnership {
  private records: StoredOwnership[] = [];
  private operation = Promise.resolve();
  private reservationSequence = 0;

  public async seed(annotations: readonly AnnotationOwnershipEntry[]): Promise<void> {
    await this.transact(() => {
      if (annotations.length > MAX_OWNED_ANNOTATIONS || !annotations.every(validAnnotation)) {
        throw new Error("Invalid memory annotation ownership fixture");
      }
      this.records = annotations.map((annotation) => ({ ...annotation }));
    });
  }

  public async load(): Promise<readonly AnnotationOwnershipEntry[]> {
    return this.transact(() => entries(this.records));
  }

  public async reserve(): Promise<string | undefined> {
    return this.transact(() => {
      const now = Date.now();
      this.records = withoutExpiredReservations(this.records, now);
      const id = this.records.length >= MAX_OWNED_ANNOTATIONS
        ? undefined
        : `reservation:memory-${++this.reservationSequence}`;
      if (!id) return undefined;
      this.records.push({ reservationId: id, expiresAt: now + RESERVATION_LIFETIME_MS });
      return id;
    });
  }

  public async commit(reservedId: string, annotationId: string): Promise<boolean> {
    return this.transact(() => {
      const result = this.commitReservation(withoutExpiredReservations(this.records, Date.now()), reservedId, annotationId);
      this.records = result.records;
      return result.committed;
    });
  }

  public async cancel(reservedId: string): Promise<void> {
    await this.transact(() => {
      this.records = this.records.filter((record) => isAnnotation(record) || record.reservationId !== reservedId);
    });
  }

  public async release(annotationId: string): Promise<void> {
    await this.transact(() => {
      this.records = this.records.filter((record) => !isAnnotation(record) || record.annotationId !== annotationId);
    });
  }

  public async recordRemoval(update: AnnotationRemovalUpdate): Promise<void> {
    await this.transact(() => {
      this.records = this.applyRemoval(this.records, update);
    });
  }

  private async transact<T>(task: () => T): Promise<T> {
    const next = this.operation.then(task, task);
    this.operation = next.then(() => undefined, () => undefined);
    return next;
  }
}

export class BrowserAnnotationOwnership extends TransactionalAnnotationOwnership {
  private static readonly key = "verseform.vfw010.annotation-ownership.v1";
  private static readonly lockName = "verseform.vfw010.annotation-ownership";

  public async load(): Promise<readonly AnnotationOwnershipEntry[]> {
    return this.transaction((records) => entries(records));
  }

  public async reserve(): Promise<string | undefined> {
    return this.transaction((records) => {
      const now = Date.now();
      const current = withoutExpiredReservations(records, now);
      const id = this.nextReservation(current);
      if (!id) return { value: undefined, records: current };
      return {
        value: id,
        records: [...current, { reservationId: id, expiresAt: now + RESERVATION_LIFETIME_MS }],
      };
    });
  }

  public async commit(reservedId: string, annotationId: string): Promise<boolean> {
    return this.transaction((records) => {
      const result = this.commitReservation(withoutExpiredReservations(records, Date.now()), reservedId, annotationId);
      return { value: result.committed, records: result.records };
    });
  }

  public async cancel(reservedId: string): Promise<void> {
    await this.transaction((records) => ({
      value: undefined,
      records: records.filter((record) => isAnnotation(record) || record.reservationId !== reservedId),
    }));
  }

  public async release(annotationId: string): Promise<void> {
    await this.transaction((records) => ({
      value: undefined,
      records: records.filter((record) => !isAnnotation(record) || record.annotationId !== annotationId),
    }));
  }

  public async recordRemoval(update: AnnotationRemovalUpdate): Promise<void> {
    await this.transaction((records) => ({
      value: undefined,
      records: this.applyRemoval(records, update),
    }));
  }

  private async transaction<T>(
    task: (records: readonly StoredOwnership[]) => T | { value: T; records: StoredOwnership[] },
  ): Promise<T> {
    const locks = globalThis.navigator?.locks;
    if (!locks || typeof locks.request !== "function") {
      throw new Error("Cross-pane annotation ownership locking is unavailable");
    }
    return locks.request(BrowserAnnotationOwnership.lockName, { mode: "exclusive" }, async () => {
      const raw = globalThis.localStorage.getItem(BrowserAnnotationOwnership.key);
      const current = raw ? strictStoredOwnership(JSON.parse(raw)) : [];
      const result = task(current);
      if (typeof result === "object" && result !== null && "records" in result && "value" in result) {
        globalThis.localStorage.setItem(BrowserAnnotationOwnership.key, JSON.stringify(result.records));
        return result.value;
      }
      return result as T;
    });
  }
}
