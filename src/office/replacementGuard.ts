/**
 * Word annotations can be accepted or rejected by host UI. Verseform only
 * replaces a critique that is still in its original Created state.
 */
export function isReplaceableAnnotationState(state: string | undefined): boolean {
  return state === "Created";
}
