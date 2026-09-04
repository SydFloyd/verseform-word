import type { Vfw010State } from "../app/controller";

export type TaskPaneView = {
  statusKind: "ready" | "blocked" | "browser";
  title: string;
  detail: string;
  preview: {
    hidden: boolean;
    heading: string;
    text: string;
  };
  insert: {
    disabled: boolean;
    label: string;
  };
  cancel: {
    disabled: boolean;
  };
  focusTarget?: "status";
};

export function taskPaneView(state: Readonly<Vfw010State>): TaskPaneView {
  return {
    statusKind: state.phase === "blocked" ? "blocked" : "ready",
    title: state.title,
    detail: state.detail,
    preview: state.preview
      ? { hidden: false, heading: state.preview.heading, text: state.preview.text }
      : { hidden: true, heading: "", text: "" },
    insert: {
      disabled: !state.canInsert,
      label: state.phase === "inserting" ? "Checking reference…" : "Insert test passage",
    },
    cancel: { disabled: !state.canCancel },
    focusTarget: state.focusTarget,
  };
}
