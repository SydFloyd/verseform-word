import "./styles.css";
import { inspectWordHost } from "./office/capabilities";

const title = document.querySelector<HTMLElement>("#status-title");
const detail = document.querySelector<HTMLElement>("#status-detail");
const card = document.querySelector<HTMLElement>(".status-card");

function renderStatus(kind: "ready" | "blocked" | "browser", heading: string, message: string): void {
  if (!title || !detail || !card) return;
  title.textContent = heading;
  detail.textContent = message;
  card.dataset.kind = kind;
}
if (typeof Office === "undefined") {
  renderStatus(
    "browser",
    "Browser preview",
    "Open this task pane inside Word to test host capabilities.",
  );
} else {
  Office.onReady(() => {
    const capability = inspectWordHost();
    renderStatus(capability.kind, capability.title, capability.detail);
  });
}
