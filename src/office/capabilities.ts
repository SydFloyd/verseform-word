export type HostCapability = {
  kind: "ready" | "blocked";
  title: string;
  detail: string;
};

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

  return {
    kind: "ready",
    title: "Word is ready",
    detail: "This host supports the annotation events required for the first walking slice.",
  };
}
