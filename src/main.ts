import "./styles.css";
import { BrowserAnnotationOwnership } from "./app/annotationOwnership";
import { Vfw010Controller } from "./app/controller";
import { inspectWordHost } from "./office/capabilities";
import { OfficeWordGateway } from "./office/wordGateway";
import { mountTaskPane } from "./ui/taskPane";

const taskPane = mountTaskPane();
if (typeof Office === "undefined") {
  taskPane.renderHostStatus(
    "browser",
    "Browser preview",
    "Open this task pane inside Word to test host capabilities.",
  );
} else {
  Office.onReady(() => {
    const capability = inspectWordHost();
    if (capability.kind === "blocked") {
      taskPane.renderHostStatus("blocked", capability.title, capability.detail);
      return;
    }
    const controller = new Vfw010Controller(
      new OfficeWordGateway(),
      taskPane.render,
      new BrowserAnnotationOwnership(),
    );
    taskPane.onInsert(async () => controller.insertSelected());
    taskPane.onCancel(async () => controller.clearSelection());
    window.addEventListener("pagehide", () => {
      void controller.stop();
    }, { once: true });
    void controller.start();
  });
}
