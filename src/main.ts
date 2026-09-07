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
let controller: VerseformController | undefined;
let startup: Promise<void> | undefined;

async function startInsideWord(): Promise<void> {
  const capability = inspectWordHost();
  if (capability.kind === "blocked") {
    taskPane.renderHostStatus("blocked", capability.title, capability.detail);
    return;
  }
  if (controller) return;

  controller = new VerseformController(
    new OfficeWordGateway(),
    taskPane.render,
    new BrowserAnnotationOwnership(),
    new DbsScriptureProvider(
      new FetchDbsTransport(),
      new BrowserScriptureCache(window.localStorage),
    ),
    new BrowserTranslationPreference(window.localStorage),
    {
      previewOnHover: document.visibilityState === "visible",
      onInteractionNeedsAttention: () => {
        void Office.addin.showAsTaskpane().catch(() => undefined);
      },
    },
  );
  taskPane.onInsert(async () => { await controller?.insertSelected(); });
  taskPane.onCancel(async () => { await controller?.clearSelection(); });
  taskPane.onTranslationChange(async (translationId) => { await controller?.selectTranslation(translationId); });
  taskPane.onClearCache(async () => { await controller?.clearScriptureCache(); });
  await controller.start();
  try {
    await Office.addin.onVisibilityModeChanged((event) => {
      controller?.setPreviewOnHover(event.visibilityMode === Office.VisibilityMode.taskpane);
    });
  } catch {
    // Pane visibility is an optional preview concern. It must never prevent the
    // shared runtime from detecting or inserting references.
    controller.setPreviewOnHover(false);
  }
}

function ensureStarted(): Promise<void> {
  startup ??= Office.onReady().then(async () => startInsideWord());
  return startup;
}

function enableVerseform(event?: Office.AddinCommands.Event): void {
  // This is a long shared runtime, so completing the queued ribbon command
  // does not tear it down. Release Word's command queue immediately; local
  // event binding, the startup preference, and DBS catalog loading continue
  // in the one document-scoped runtime without making repeated clicks useful.
  event?.completed();
  void ensureStarted()
    .then(async () => {
      controller?.setPreviewOnHover(false);
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    })
    .catch(() => {
      const capability = officeReadinessFailure();
      taskPane.renderHostStatus("blocked", capability.title, capability.detail);
    });
}

function fillScripture(event?: Office.AddinCommands.Event): void {
  // The command is the explicit mutation intent. Keep Word's command context
  // alive until the selection has been captured and the guarded operation has
  // finished; completing it first can discard the document selection.
  void ensureStarted()
    .then(async () => { await controller?.fillAtSelection(); })
    .catch(() => {
      const capability = officeReadinessFailure();
      taskPane.renderHostStatus("blocked", capability.title, capability.detail);
      void Office.addin.showAsTaskpane().catch(() => undefined);
    })
    .finally(() => { event?.completed(); });
}

if (typeof Office === "undefined") {
  taskPane.renderHostStatus(
    "browser",
    "Browser preview",
    "Open this task pane inside Word to test host capabilities.",
  );
} else {
  Office.actions.associate("enableVerseform", enableVerseform);
  Office.actions.associate("fillScripture", fillScripture);
  void ensureStarted().catch(() => {
    const capability = officeReadinessFailure();
    taskPane.renderHostStatus("blocked", capability.title, capability.detail);
  });
}
