import type { VerseformState } from "../app/controller";
import { taskPaneView } from "./viewModel";

export type MountedTaskPane = {
  render(state: Readonly<VerseformState>): void;
  onInsert(action: () => Promise<void>): void;
  onCancel(action: () => Promise<void>): void;
  onTranslationChange(action: (translationId: string) => Promise<void>): void;
  onClearCache(action: () => Promise<void>): void;
  renderHostStatus(kind: "blocked" | "browser", title: string, detail: string): void;
};

function required<T extends HTMLElement>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) throw new Error(`Task-pane element is missing: ${selector}`);
  return element;
}

export function mountTaskPane(): MountedTaskPane {
  const card = required<HTMLElement>(".status-card");
  const title = required<HTMLElement>("#status-title");
  const detail = required<HTMLElement>("#status-detail");
  const preview = required<HTMLElement>("#preview-card");
  const previewHeading = required<HTMLElement>("#preview-heading");
  const previewText = required<HTMLElement>("#preview-text");
  const previewAttribution = required<HTMLElement>("#preview-attribution");
  const insert = required<HTMLButtonElement>("#insert-button");
  const cancel = required<HTMLButtonElement>("#cancel-button");
  const translation = required<HTMLSelectElement>("#translation-select");
  const translationNotice = required<HTMLElement>("#translation-notice");
  const clearCache = required<HTMLButtonElement>("#clear-cache-button");
  let insertAction: (() => Promise<void>) | undefined;
  let cancelAction: (() => Promise<void>) | undefined;
  let translationAction: ((translationId: string) => Promise<void>) | undefined;
  let clearCacheAction: (() => Promise<void>) | undefined;
  let translationSignature = "";

  const actionFailure = (): void => {
    card.dataset.kind = "blocked";
    title.textContent = "Verseform could not complete that action";
    detail.textContent = "Verseform could not confirm the result. Inspect the document before retrying or reopening the task pane.";
    preview.hidden = true;
    insert.disabled = true;
    cancel.disabled = true;
    translation.disabled = true;
    clearCache.disabled = true;
    title.focus();
  };

  insert.addEventListener("click", () => {
    if (!insertAction || insert.disabled) return;
    void insertAction().catch(actionFailure);
  });
  cancel.addEventListener("click", () => {
    if (!cancelAction || cancel.disabled) return;
    void cancelAction().catch(actionFailure);
  });
  translation.addEventListener("change", () => {
    if (!translationAction || translation.disabled || !translation.value) return;
    void translationAction(translation.value).catch(actionFailure);
  });
  clearCache.addEventListener("click", () => {
    if (!clearCacheAction || clearCache.disabled) return;
    void clearCacheAction().catch(actionFailure);
  });

  return {
    render(state) {
      const view = taskPaneView(state);
      card.dataset.kind = view.statusKind;
      title.textContent = view.title;
      detail.textContent = view.detail;
      preview.hidden = view.preview.hidden;
      previewHeading.textContent = view.preview.heading;
      previewText.textContent = view.preview.text;
      previewAttribution.textContent = view.preview.attribution;
      insert.disabled = view.insert.disabled;
      insert.textContent = view.insert.label;
      cancel.disabled = view.cancel.disabled;
      const nextSignature = view.translation.options.map((option) => `${option.id}:${option.label}`).join("|");
      if (translationSignature !== nextSignature) {
        const options = view.translation.options.length
          ? view.translation.options.map((item) => {
            const option = document.createElement("option");
            option.value = item.id;
            option.textContent = item.label;
            return option;
          })
          : [Object.assign(document.createElement("option"), {
            value: "",
            textContent: "Translations unavailable",
          })];
        translation.replaceChildren(...options);
        translationSignature = nextSignature;
      }
      translation.value = view.translation.selectedId;
      translation.disabled = view.translation.disabled;
      translationNotice.textContent = view.translation.notice;
      clearCache.disabled = view.clearCache.disabled;
      if (view.focusTarget === "status") title.focus();
    },
    onInsert(action) {
      insertAction = action;
    },
    onCancel(action) {
      cancelAction = action;
    },
    onTranslationChange(action) {
      translationAction = action;
    },
    onClearCache(action) {
      clearCacheAction = action;
    },
    renderHostStatus(kind, hostTitle, hostDetail) {
      card.dataset.kind = kind;
      title.textContent = hostTitle;
      detail.textContent = hostDetail;
      preview.hidden = true;
      insert.disabled = true;
      cancel.disabled = true;
      translation.disabled = true;
      clearCache.disabled = true;
      title.focus();
    },
  };
}
