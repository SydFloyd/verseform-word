import type { Vfw010State } from "../app/controller";
import { taskPaneView } from "./viewModel";

export type MountedTaskPane = {
  render(state: Readonly<Vfw010State>): void;
  onInsert(action: () => Promise<void>): void;
  onCancel(action: () => Promise<void>): void;
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
  const insert = required<HTMLButtonElement>("#insert-button");
  const cancel = required<HTMLButtonElement>("#cancel-button");
  let insertAction: (() => Promise<void>) | undefined;
  let cancelAction: (() => Promise<void>) | undefined;

  const actionFailure = (): void => {
    card.dataset.kind = "blocked";
    title.textContent = "Verseform could not complete that action";
    detail.textContent = "Verseform could not confirm the result. Inspect the document before retrying or reopening the task pane.";
    preview.hidden = true;
    insert.disabled = true;
    cancel.disabled = true;
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

  return {
    render(state) {
      const view = taskPaneView(state);
      card.dataset.kind = view.statusKind;
      title.textContent = view.title;
      detail.textContent = view.detail;
      preview.hidden = view.preview.hidden;
      previewHeading.textContent = view.preview.heading;
      previewText.textContent = view.preview.text;
      insert.disabled = view.insert.disabled;
      insert.textContent = view.insert.label;
      cancel.disabled = view.cancel.disabled;
      if (view.focusTarget === "status") title.focus();
    },
    onInsert(action) {
      insertAction = action;
    },
    onCancel(action) {
      cancelAction = action;
    },
    renderHostStatus(kind, hostTitle, hostDetail) {
      card.dataset.kind = kind;
      title.textContent = hostTitle;
      detail.textContent = hostDetail;
      preview.hidden = true;
      insert.disabled = true;
      cancel.disabled = true;
      title.focus();
    },
  };
}
