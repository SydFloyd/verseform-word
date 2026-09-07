# Verseform for Word — Product Requirements

## Purpose

Verseform for Word helps a person remain inside Microsoft Word while inserting accurate, authorized Scripture. The add-in should feel like a small native writing aid, not a second editor or Bible study suite.

## Central interaction

1. The user types a supported reference and leaves the cursor in or immediately after it. **Fill Scripture** starts the shared runtime when needed and inserts without requiring an open task pane or a separate enable step. Opening **Preview & Settings** also starts the same runtime and provides optional preview, translation, cache, privacy, licensing, and help controls.
2. After the user completes a supported reference and types a delimiter, Verseform scans only the relevant paragraph locally.
3. A valid reference receives a temporary Word-native annotation. Invalid coordinates may receive restrained, explanatory feedback; ordinary prose stays untouched.
4. Hover is passive while the pane is hidden. If the user intentionally opens **Preview & Settings**, hover may preview the exact reference there when Word delivers its best-effort hover event; it never changes prose. The required pane-free insertion action is Word's **Fill Scripture** ribbon command or selected-text context-menu command. It targets exactly one reference containing or immediately preceding the cursor; choosing the command is itself a completion boundary when the reference ends the paragraph. Annotation click or `Alt+Down` may use the same guarded path when the host delivers that event. If an explicit request fails while the pane is closed, Verseform opens the optional pane to show the failure instead of failing invisibly.
5. The explicit Fill, annotation, or pane action replaces exactly that resolved reference with passage text followed by a compact editable citation after every freshness check passes. It does not append a full copyright paragraph after every quotation. The optional pane retains the selected translation's full provider notice for document-level attribution use and its **Insert passage** control remains an accessible alternative. Verseform does not use WordApi 1.8 critique suggestions as command buttons because Word owns suggestion acceptance as a document replacement and it cannot substitute for Verseform's guarded range command.
6. Before replacement, Verseform rereads and verifies that the paragraph, annotation-owned source range, reference, and translation still match the request. It refuses every stale, missing, ambiguous, or failed state observable through WordApi 1.7 and explains that the writing changed. WordApi 1.7 has no atomic conditional replace, so a coauthor edit after that final verification and before Word applies the queued replace cannot be ruled out; this narrow platform limitation is documented rather than claimed safe.
7. The replacement is one Word transaction so Word's ordinary Undo restores the reference.

## Reference intelligence

- The initial language is English. Full names, approved abbreviations, conservative fuzzy book-name matching, single verses, and same-chapter ranges follow Verseform's proven behavior.
- Chapter and verse bounds are strict. Only book names may be fuzzy.
- One request may contain a single verse or a same-chapter range of at most 25 verses. Longer ranges are refused locally with an explanation and never reach DBS; the writer may split them into smaller references.
- Whitespace, punctuation, user-created paragraph breaks, and common closing characters count as delimiters. Because Word omits its paragraph mark from `Paragraph.text`, the Word boundary must prove the real break rather than treating every current paragraph end as a delimiter. An unfinished reference is not annotated.
- URLs, emails, existing generated citations, and representative prose false positives are excluded.
- Detection never contacts Word services beyond reading the changed paragraph and never contacts DBS.
- Multilingual detection is deferred until DBS has tested and approved the product direction and the English pilot user has supplied practical feedback. Any later language remains a separate capability requiring its own corpus-backed ambiguity, bounds, and Word offset proof; no current behavior claims another language.

## Scripture and translations

- Verseform lists translations authorized by the DBS ARC API and remembers a local preference.
- If there is no preference, NASB is preferred when the catalog contains it; otherwise the first clearly available authorized translation is used.
- No WEB or other Bible is bundled. When DBS is unavailable, detection can continue but preview and insertion report that Scripture text needs a connection.
- Provider requests contain only translation ID, canonical book ID, and chapter. Document prose, file name, title, identity, and surrounding words are never sent.
- Responses are time-, size-, and schema-bounded, normalized as untrusted plain text, and never executed or rendered as provider HTML.
- Authorized catalog and Scripture responses may be cached locally. Persistent caches must be schema-versioned, bounded by age, bytes, and entries, clearable by the user, and must never contain document prose or identity.
- Inserted text includes the normalized passage and an editable citation with translation abbreviation. The optional pane exposes the selected translation's complete provider copyright notice so the writer can apply document-level notice requirements without duplicating it after every quotation.

## Word behavior

- One Office.js add-in with a long SharedRuntime 1.1 lifetime targets Word on the web and supported Word desktop clients. The same core and manifest contract must be used across hosts.
- Native Word APIs own annotations, selection, replacement, formatting inheritance, Undo, coauthor behavior, and accessibility semantics.
- Temporary annotations must not persist as document content or alter saved prose.
- Closing or hiding the optional pane must not stop an active shared runtime, lose its event handlers, or require the writer to keep a side panel open. A pane-document `pagehide` or unload signal is not proof that the long shared runtime ended and must never invoke controller teardown. Fill and Preview & Settings start the runtime idempotently when needed; no separate enable state or load-on-reopen preference is exposed, and a preference saved by an earlier release is retired without blocking startup. The add-in must still behave safely when the runtime is actually destroyed, coauthors change text, the selection moves, or the host lacks WordApi 1.7 or SharedRuntime 1.1.
- WordApi 1.7 annotation click, hover, and removal events are not coauthor-triggered. Verseform must not present those events, or its final reread, as an atomic coauthor stale guarantee.
- Word on the web and Word for Windows must expose **Fill Scripture** on the ribbon and selected-text context menu while the pane is closed. The command must resolve locally to exactly one reference, refuse an ambiguous or changed target, request only normalized DBS coordinates, and replace through the same final freshness and one-Undo path. Annotation click and `Alt+Down` remain supported conveniences where Word delivers them, but they are not the pane-free release gate. Hover preview is an optional visible-pane enhancement.
- WordApi 1.7 annotations require a connected Microsoft 365 subscription. This limitation is stated before installation and inside an unsupported host.

## Privacy, accessibility, and distribution

- No Verseform account, advertising, analytics, telemetry, crash upload, AI prose processing, or background scan exists.
- The production add-in is served over HTTPS from an owner-approved host. Hosting logs and retention must be disclosed.
- Keyboard and screen-reader users can reach Word's Fill command and the task pane, understand status, preview a reference, insert it, cancel, and recover focus. Do not advertise a custom simultaneous shortcut until its registration is proved in the release hosts.
- The task pane must declare and paint a matching light or dark canvas/text surface, and use paired system colors when Windows forced-colors mode is active; it must never rely on the host canvas for readable text.
- DBS is credited clearly and no endorsement by Microsoft or DBS is implied without written approval.
- Public distribution must use a production-supported manifest and pass Microsoft validation. The current unified Word manifest is not a production dependency while Microsoft labels it preview.

## Initial exclusions

Bundled/offline Bibles, whole-document automatic replacement, Google Docs, Outlook, PowerPoint, VBA/VSTO, mobile-specific UI, AI writing, document upload, accounts, sync, collaboration features, bibliography management, and general Word formatting tools are outside the initial product.

## Acceptance

The first usable release is ready when a Word user can type a supported reference and invoke Fill without enabling or opening a persistent pane; replace only that exact reference with passage and compact citation after its final observable freshness check; Undo back to the reference; and receive clear failure behavior for observable stale text, unavailable DBS, invalid or overlong references, unsupported hosts, and a destroyed/reloaded runtime. After either Fill or Preview & Settings starts the shared runtime, closing the optional pane must not stop local detection. That pane retains preview, Insert, translation notice, cache, privacy, licensing, and concise usage/limit help, and the documented WordApi 1.7 check-then-replace coauthor timing limitation remains explicit.
