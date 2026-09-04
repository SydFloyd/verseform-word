import { isTranslationId, type TranslationPreferenceStore } from "../core/scripture";

export const TRANSLATION_PREFERENCE_KEY = "verseform-word.preferred-translation.v1";
type StorageLike = Pick<Storage, "getItem" | "setItem">;

export class BrowserTranslationPreference implements TranslationPreferenceStore {
  public constructor(private readonly storage: StorageLike) {}

  public async load(): Promise<string | undefined> {
    try {
      const value = this.storage.getItem(TRANSLATION_PREFERENCE_KEY);
      return value && isTranslationId(value) ? value : undefined;
    } catch {
      return undefined;
    }
  }

  public async save(translationId: string): Promise<void> {
    if (!isTranslationId(translationId)) throw new Error("That translation identifier is invalid.");
    try {
      this.storage.setItem(TRANSLATION_PREFERENCE_KEY, translationId);
    } catch {
      throw new Error("The translation preference could not be saved on this device.");
    }
  }
}

export class MemoryTranslationPreference implements TranslationPreferenceStore {
  public constructor(private value?: string) {}
  public async load(): Promise<string | undefined> { return this.value; }
  public async save(translationId: string): Promise<void> { this.value = translationId; }
}
