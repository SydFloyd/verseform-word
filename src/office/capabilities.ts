export type HostCapability = {
  kind: "ready" | "blocked";
  title: string;
  detail: string;
};

/** A rejected Office readiness promise has not reached the Word boundary. */
export function officeReadinessFailure(): HostCapability {
  return {
    kind: "blocked",
    title: "Word could not start Verseform",
    detail: "Verseform did not change your document. Reopen the task pane inside a connected Microsoft Word host with WordApi 1.7.",
  };
}

export function inspectWordHost(): HostCapability {
  if (Office.context.host !== Office.HostType.Word) {
    return {
      kind: "blocked",
      title: "Open in Word",
      detail: "This add-in works only inside Microsoft Word.",
    };
  }

  if (!Office.context.requirements.isSetSupported("WordApi", "1.7")) {
    return {
      kind: "blocked",
      title: "Word needs annotation support",
      detail: "Verseform requires WordApi 1.7 and a connected Microsoft 365 subscription.",
    };
  }

  if (!Office.context.requirements.isSetSupported("SharedRuntime", "1.1")) {
    return {
      kind: "blocked",
      title: "Word needs background add-in support",
      detail: "Verseform requires SharedRuntime 1.1 so detection can continue while its optional pane is closed.",
    };
  }

  return {
    kind: "ready",
    title: "Word is ready",
    detail: "Verseform can run for this document while its optional preview and settings pane is closed.",
  };
}
