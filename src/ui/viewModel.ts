import type { VerseformState } from "../app/controller";
import { VFW010_TRANSLATION } from "../app/interaction";

export type TaskPaneView = {
  statusKind: "ready" | "blocked" | "browser";
  title: string;
  detail: string;
  preview: {
    hidden: boolean;
    heading: string;
    text: string;
    attribution: string;
  };
  insert: {
    disabled: boolean;
    label: string;
  };
  cancel: {
    disabled: boolean;
  };
  translation: {
    disabled: boolean;
    selectedId: string;
    options: Array<{ id: string; label: string }>;
    notice: string;
  };
  clearCache: { disabled: boolean };
  focusTarget?: "status";
};

export function taskPaneView(state: Readonly<VerseformState>): TaskPaneView {
  const selectedTranslation = state.translations?.find(
    (translation) => translation.id === state.selectedTranslationId,
  );
  return {
    statusKind: state.phase === "blocked" ? "blocked" : "ready",
    title: state.title,
    detail: state.detail,
    preview: state.preview
      ? {
        hidden: false,
        heading: state.preview.heading,
        text: state.preview.text,
        attribution: state.preview.attribution,
      }
      : { hidden: true, heading: "", text: "", attribution: "" },
    insert: {
      disabled: !state.canInsert,
      label: state.phase === "inserting"
        ? "Checking reference…"
        : state.preview?.translationId === VFW010_TRANSLATION.id
          ? "Insert test passage"
          : state.preview
            ? `Insert ${state.preview.citationLabel}`
            : "Insert passage",
    },
    cancel: { disabled: !state.canCancel },
    translation: {
      disabled: !state.canSelectTranslation,
      selectedId: state.selectedTranslationId ?? "",
      options: (state.translations ?? []).map((translation) => ({
        id: translation.id,
        label: `${translation.citationLabel} — ${translation.name}`,
      })),
      notice: selectedTranslation?.attribution ?? "Translation copyright notice loads with the authorized DBS catalog.",
    },
    clearCache: { disabled: !state.canClearCache },
    focusTarget: state.focusTarget,
  };
}
