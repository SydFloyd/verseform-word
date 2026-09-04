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
- **annotations:** annotation ID → paragraph ID, source text, normalized reference, source range, paragraph revision, and generation.
- **scripture:** catalog, preferred/effective translation, request stamp, bounded preview result, and failure.
- **ui:** selected annotation, task-pane phase, and actionable status.

Events describe facts; effects request work. Each asynchronous effect carries a monotonically increasing generation plus the paragraph revision and translation ID it depends on. Results are ignored unless every stamp still matches.

## Word boundary

WordApi 1.6 provides paragraph-change events and stable-per-session paragraph IDs. WordApi 1.7 provides temporary critique annotations and hover, click, insertion, removal, and popup-action events. Event handlers exist only while the add-in runtime is alive and must be registered again after reopening.

VFW-010 must prove four uncertain host facts in Word on the web rather than infer them from types:

1. delimiter typing yields a changed-paragraph event soon enough for calm feedback;
2. annotation offsets align with the paragraph text returned by Word;
3. hover/click/popup actions can identify the exact annotation without mutating prose;
4. one explicit action can replace that exact occurrence with ordinary Word Undo support, including duplicate references in one paragraph.

If the native popup cannot safely provide a live passage, hover selects the annotation and the task pane provides the preview and Insert action. If exact duplicate replacement cannot be proved with stable APIs, the add-in must refuse ambiguity or require a current selection; it must not guess.

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

## Proof economy

- Pure tests own parser, canon, false positives, delimiters, freshness, and insertion text.
- Fake ports own cancellation, stale results, translation changes, duplicate annotations, and provider failures.
- Browser tests own the task-pane state and accessibility without pretending to be Word.
- Word on the web owns annotation offsets/events, exact replacement, Undo, task-pane lifecycle, and subscription behavior.
- Word for Windows repeats only the cross-host contract after the online flow is stable.

Every defect leaves a regression at the cheapest layer that can reproduce it. `npm run check` is the repository gate and includes local structural manifest validation. Microsoft's external validator remains a release contract gate rather than a vulnerable build dependency.
