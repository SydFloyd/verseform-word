# Verseform for Word — System Design

## System shape

Verseform for Word is a pure reference-and-freshness core inside a narrow Office.js shell. Word owns the document canvas. DBS owns authorized Scripture delivery. Verseform coordinates them without copying the document into a second editor or allowing either boundary to define product rules.

```text
Word changed paragraph / annotation event
                │
                ▼
        Office.js Word gateway
                │  paragraph snapshot or user intent
                ▼
       small application kernel
         │                 │
         ▼                 ▼
 pure reference core    DBS scripture port
         │                 │
         └──── plan/result ┘
                │
                ▼
 temporary annotation / task-pane preview / guarded Word replacement
```

## Module ownership

```text
src/core/       Pure canon, parser, normalized reference, freshness, insertion plan
src/app/        Closed state/event/effect model, controller, ports, selectors
src/office/     Office.js capability, paragraph, annotation, range, and undo gateway
src/adapters/   DBS JSON transport and local preference/cache implementations
src/ui/         Small task-pane projection of the application view model
tests/          Pure corpus, fake Word/DBS contracts, browser task-pane cases
manifest.xml    Word activation, permission, command, resource, and API requirements
```

A module exists only when it owns a rule or an external adaptation. React and a state framework are unnecessary until the task pane proves complex enough to justify them.

## State tower

The kernel owns one canonical state with five regions:

- **host:** API support, handler generation, active/closed state, last host error.
- **paragraphs:** revision and redacted structural identity for paragraphs currently carrying Verseform annotations; never a persisted copy of document prose.
- **annotations:** annotation ID → paragraph ID, source text, normalized reference, source range, and paragraph revision; the controller owns the enclosing runtime generation.
- **scripture:** catalog, preferred/effective translation, request stamp, bounded preview result, and failure.
- **ui:** selected annotation, task-pane phase, fake host-proof preview, and actionable status.

Events describe facts; effects request work. Each task-pane runtime binds handlers under a monotonically increasing generation and returns a stop handle that removes those exact handlers. Delayed events from a stopped generation are ignored. Paragraph revisions and translation IDs stamp annotation actions; future provider effects also carry their own request generation. Results are ignored unless every applicable stamp still matches.

## Word boundary

WordApi 1.6 provides paragraph-change events and stable-per-session paragraph IDs. WordApi 1.7 provides temporary critique annotations plus hover, click, insertion, and removal events. The installed Office.js type contract marks annotation popup-action events and popup options as WordApi 1.8, so VFW-010 uses an accessible task-pane preview and Insert action instead. Event handlers exist only while the add-in runtime is alive and must be registered again after reopening.

The Office gateway retains the event-registration results and an original-context document anchor in a `WordRuntime` stop handle. Startup attempts to clean registrations that were queued before a failure; teardown asks Word to remove every registered handler through that retained context and asks it to clear each controller-owned temporary annotation. Annotation cleanup reports individual failures. While the pane remains alive, a failed cleanup retains ownership, blocks new annotations in that paragraph, and can retry on a later paragraph event; stale or rejected insertion instead reports the cleanup failure and never claims the marker was removed.

`pagehide` cannot prove that an async teardown reaches Word. Therefore VFW-010 keeps a bounded, app-local ownership ledger of at most 256 opaque annotation IDs. Browser ownership mutations use a same-origin `navigator.locks` exclusive transaction that rereads current storage and applies only a delta; the lock never spans a Word call. If locking is unavailable, startup blocks before it binds Word events. Before Word creates a marker, Verseform atomically reserves one bounded opaque slot; after Word returns the ID, it atomically commits that ID before the controller publishes it as actionable. A failed commit triggers immediate exact-ID host cleanup. If the ledger is full, no host annotation is created. Reservations contain only a random opaque token and one-minute expiry, so a crashed pre-create reservation cannot consume capacity forever. A reservation is not an annotation owner record: a crash after Word creates a marker but before the returned ID commits remains an explicitly unrecoverable WordApi 1.7 gap. On every fresh pane runtime, Verseform asks the current document to remove exactly the stored IDs before binding new handlers. Even a successful startup delete keeps its opaque ID, because Word Undo can return it; only a later startup result that says the ID is missing increments its bounded retry count. A missing ID is no longer actionable in that runtime, but is retained for up to eight current-document misses: it may belong to another open document and must be recoverable when that document reopens. A missing sweep never blocks startup. Stop cleans both active and retired IDs from the current runtime without aging them, so closing an unrelated document does not age another document's retained entry. It never enumerates or removes annotations by color, title, range, or nearby prose, and it refuses an unknown ID rather than guessing.

Word on the web has now shown that Undo can resurrect every native annotation from the paragraph state before a replacement—not just the selected annotation—even after the controller asked Word to clear one during its changed-paragraph refresh. While one pane runtime is alive, the controller moves every active annotation affected by a paragraph refresh, replacement, stale rejection, or annotation-removed event into a bounded runtime-only retired map before cleanup. It keeps that opaque ID in the ledger even when the current cleanup says removed or missing. A candidate or a late activation can reactivate it only after `getAnnotationById` proves the same ID is `Created`, in the same paragraph, at the same UTF-16 offset and length, with the same source text. If a late exact resurrection overlaps a newer controller-owned marker, the controller retires and exact-deletes only that newer known ID; cleanup failure blocks the action. The fake-host model proves this same-runtime recovery path, not that Word on the web has delivered every event ordering safely. The controller does not evict a retired ID: capacity exhaustion blocks new markers rather than making an ID Word may resurrect unowned. The retired map contains paragraph/source metadata only in the live task-pane process; it never enters browser storage.

The ledger contains no paragraph text, document name, document identity, account data, or Scripture data. Its browser representation is parsed strictly: malformed roots or entries, duplicate IDs, and overflow leave the existing record untouched and block Word startup rather than silently erasing recoverable ownership. The lock and browser-storage assumptions are local fake/browser-contract evidence only; Word on the web and Windows must still prove them in the intended task-pane engines. On pane close, the retired map is deliberately forgotten because persisting its paragraph/source data would violate the local-only metadata boundary. Startup can exact-sweep a stored ID that is present then, but cannot safely map an ID that Word resurrects by Undo after reopening; bounded missing-ID expiry has the same limitation. WordApi 1.7 supplies neither a durable owner tag nor a safe way to distinguish all add-in annotations, so it cannot repair that case by scanning or deleting on appearance. This remains a VFW-010 feasibility gate, not a completed Undo/lifecycle claim. It also does not expose an atomic Word-create-plus-local-commit operation: a crash after host annotation creation but before commit, a corrupted/local-storage collision, or simultaneous commit and host-cleanup failure can still leave an unowned marker that this add-in must not sweep. The reservation and exact-ID retry narrow that window but do not prove it absent. SharedRuntime can keep code running while a pane is hidden, but Microsoft’s documentation only explicitly promises `VisibilityModeChanged` for programmatic hide/show and says hidden listeners continue running; until native-close behavior is observed, VFW-010 does not use it to risk background scanning. These are implementation guards, not evidence that Word on the web will deliver or destroy events in a particular order; VFW-010's host run still owns those claims.

The Word on the web walk has proved these host facts rather than inferring them from types:

1. delimiter typing yields a changed-paragraph event soon enough for calm feedback;
2. annotation offsets align with the paragraph text returned by Word;
3. click and keyboard activation can identify the exact annotation without mutating prose; and
4. one explicit action can replace that exact occurrence with ordinary Word Undo support, including duplicate references in one paragraph.

VFW-010 remains open until bounded host proof covers hover identity, screen-reader and forced-color operation, unsupported-host messaging, cross-document and close-interruption ledger behavior, distinguishable late post-close Undo resurrection, and the check-then-replace coauthor race—or the interaction changes so those claims are no longer required.

VFW-010 maps each candidate by inserting one critique at a time and recording its returned annotation ID, rather than assuming a result-array ordering for duplicates. Hover or click selects that ID and the task pane provides the preview and Insert action. During replacement, the Word adapter reloads the paragraph and the annotation's own range, source text, critique offset, critique length, and owning paragraph before queuing one `Range.insertText(..., "Replace")` mutation. It refuses unknown, missing, stale, or ambiguous mappings.

This is the strongest guard the 1.7 surface exposes locally, but it is not an atomic conditional replace: validation and the replace command require separate `context.sync()` boundaries, so a coauthor edit can theoretically land between them. Fake-host tests must not be presented as proof that this race is impossible. The VFW-010 Word on the web run must determine whether annotation invalidation/event ordering closes the race in practice; otherwise the interaction must remain a feasibility failure rather than silently claim stale-safe insertion.

## Detection boundary

The detector consumes plain paragraph text and a canon contract, returning valid and invalid candidates with exact UTF-16 offsets. It knows no Office.js, DOM, provider, translation catalog, or network.

Verseform's English parser is the initial behavioral reference because it already proves delimiter gating, strict bounds, conservative fuzziness, URL rejection, and false-positive control. DBS's MIT detector is valuable for multilingual book-name data and language patterns, but its DOM transformation, chapter-only/list grammar, direct HTML popup, and unvalidated numeric coordinates do not enter the core wholesale. Any reused DBS code or data is copied locally, audited, corpus-backed, and credited; production never loads the remote detector script.

## Scripture boundary

The DBS adapter exposes two operations: list authorized translations and fetch one normalized chapter. It validates identifiers before URL construction, uses HTTPS only, bounds time and body size, parses JSON as untrusted input, normalizes verse text, and selects the requested verse range locally.

Detection never calls this port. Hover or activation may start a cancellable request. A replacement plan is accepted only after rereading the Word paragraph and proving its recorded source slice is unchanged. No WEB fallback exists.

DBS cache permission is confirmed for this add-in. VFW-010 remains provider-free and needs no persistent Scripture cache. VFW-020 may add one local adapter that stores only authorized catalog/chapter responses after its origin, age, byte, entry, schema, clear, and removal bounds have executable proof; document prose and identity are never cache inputs.

## Manifest and hosting

The add-in-only XML manifest requests `ReadWriteDocument` and WordApi 1.7, adds one Home-ribbon command, and points to an HTTPS task pane. It is used because the unified manifest is still preview for production Word add-ins. Development uses a trusted localhost certificate; production substitutes one owner-approved HTTPS origin in all manifest resources.

Only Microsoft's required Office.js CDN and the allowlisted DBS ARC origin are runtime dependencies. The task pane CSP denies all other network destinations. Production hosting may record ordinary HTTP metadata; it must not receive document content and its retention policy must be documented before release.

The task pane declares `color-scheme: light dark` in its document and root CSS, paints its own canvas/text/surface colors for both `prefers-color-scheme` values, and uses paired CSS system colors under `forced-colors: active`. This prevents an embedded Word dark canvas from becoming an implicit text background and preserves the user’s high-contrast palette. A Word on the web dark-theme run showed readable canvas, text, controls, disabled states, and focus after this repair. Structural tests cover both schemes, but Windows forced-colors and screen-reader behavior remain host observations.

## Proof economy

- Pure tests own parser, canon, false positives, delimiters, freshness, and insertion text.
- Fake ports own cancellation, stale results, translation changes, duplicate annotations, and provider failures.
- Browser tests own the task-pane state and accessibility without pretending to be Word.
- Word on the web owns annotation offsets/events, exact replacement, Undo, task-pane lifecycle, and subscription behavior.
- Word for Windows repeats only the cross-host contract after the online flow is stable.

Every defect leaves a regression at the cheapest layer that can reproduce it. `npm run check` is the repository gate and includes local structural manifest validation. Microsoft's external validator remains a release contract gate rather than a vulnerable build dependency.
