import "./styles.css";
import { BrowserAnnotationOwnership } from "./app/annotationOwnership";
import { VerseformController } from "./app/controller";
import { DbsScriptureProvider, FetchDbsTransport } from "./adapters/dbsScriptureProvider";
import { BrowserScriptureCache } from "./adapters/scriptureCache";
import { BrowserTranslationPreference } from "./adapters/translationPreference";
import { inspectWordHost, officeReadinessFailure } from "./office/capabilities";
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
    const controller = new VerseformController(
      new OfficeWordGateway(),
      taskPane.render,
      new BrowserAnnotationOwnership(),
      new DbsScriptureProvider(
        new FetchDbsTransport(),
        new BrowserScriptureCache(window.localStorage),
      ),
      new BrowserTranslationPreference(window.localStorage),
    );
    taskPane.onInsert(async () => controller.insertSelected());
    taskPane.onCancel(async () => controller.clearSelection());
    taskPane.onTranslationChange(async (translationId) => controller.selectTranslation(translationId));
    taskPane.onClearCache(async () => controller.clearScriptureCache());
    window.addEventListener("pagehide", () => {
      void controller.stop();
    }, { once: true });
    void controller.start();
  }).catch(() => {
    const capability = officeReadinessFailure();
    taskPane.renderHostStatus("blocked", capability.title, capability.detail);
  });
}
