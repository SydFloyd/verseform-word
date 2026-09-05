# Verseform for Word — Product Requirements

## Purpose

Verseform for Word helps a person remain inside Microsoft Word while inserting accurate, authorized Scripture. The add-in should feel like a small native writing aid, not a second editor or Bible study suite.

## Central interaction

1. While the Verseform add-in is active, Word reports a changed paragraph.
2. After the user completes a supported reference and types a delimiter, Verseform scans only the relevant paragraph locally.
3. A valid reference receives a temporary Word-native annotation. Invalid coordinates may receive restrained, explanatory feedback; ordinary prose stays untouched.
4. Hover previews the exact reference in the task pane when Word delivers its hover event; hover remains a best-effort convenience and never changes prose. Clicking the annotation or selecting it with Word's `Alt+Down` command is the explicit high-throughput insertion action. It reuses a ready or in-flight hover preview when possible, or requests the passage on demand when hover was unavailable.
5. The explicit click or keyboard action replaces exactly that activated reference with passage text followed by an editable citation and attribution after every freshness check passes. The task-pane **Insert passage** control remains an accessible alternative. Verseform does not use WordApi 1.8 critique suggestions as command buttons because Word owns suggestion acceptance as a document replacement and it cannot substitute for Verseform's guarded range command.
6. Before replacement, Verseform rereads and verifies that the paragraph, annotation-owned source range, reference, and translation still match the request. It refuses every stale, missing, ambiguous, or failed state observable through WordApi 1.7 and explains that the writing changed. WordApi 1.7 has no atomic conditional replace, so a coauthor edit after that final verification and before Word applies the queued replace cannot be ruled out; this narrow platform limitation is documented rather than claimed safe.
7. The replacement is one Word transaction so Word's ordinary Undo restores the reference.

## Reference intelligence

- The initial language is English. Full names, approved abbreviations, conservative fuzzy book-name matching, single verses, and same-chapter ranges follow Verseform's proven behavior.
- Chapter and verse bounds are strict. Only book names may be fuzzy.
- Whitespace, punctuation, paragraph breaks, and common closing characters count as delimiters. An unfinished reference is not annotated.
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
- Inserted text includes the normalized passage, an editable citation with translation abbreviation, and the provider copyright/attribution as an editable visible line. The same attribution remains visible in the task pane before insertion.

## Word behavior

- One Office.js task-pane add-in targets Word on the web and supported Word desktop clients. The same core and manifest contract must be used across hosts.
- Native Word APIs own annotations, selection, replacement, formatting inheritance, Undo, coauthor behavior, and accessibility semantics.
- Temporary annotations must not persist as document content or alter saved prose.
- The add-in must behave safely when its task pane closes, event handlers are destroyed, coauthors change text, the selection moves, or the host lacks WordApi 1.7.
- WordApi 1.7 annotation click, hover, and removal events are not coauthor-triggered. Verseform must not present those events, or its final reread, as an atomic coauthor stale guarantee.
- Word on the web and Word for Windows click and `Alt+Down` insertion are required. Hover preview is progressive enhancement because a host may expose the API without delivering the event reliably; a click must still request, verify, and insert safely when no hover preview exists.
- WordApi 1.7 annotations require a connected Microsoft 365 subscription. This limitation is stated before installation and inside an unsupported host.

## Privacy, accessibility, and distribution

- No Verseform account, advertising, analytics, telemetry, crash upload, AI prose processing, or background scan exists.
- The production add-in is served over HTTPS from an owner-approved host. Hosting logs and retention must be disclosed.
- Keyboard and screen-reader users can reach the task pane, understand status, preview a reference, insert it, cancel, and recover focus.
- The task pane must declare and paint a matching light or dark canvas/text surface, and use paired system colors when Windows forced-colors mode is active; it must never rely on the host canvas for readable text.
- DBS is credited clearly and no endorsement by Microsoft or DBS is implied without written approval.
- Public distribution must use a production-supported manifest and pass Microsoft validation. The current unified Word manifest is not a production dependency while Microsoft labels it preview.

## Initial exclusions

Bundled/offline Bibles, whole-document automatic replacement, Google Docs, Outlook, PowerPoint, VBA/VSTO, mobile-specific UI, AI writing, document upload, accounts, sync, collaboration features, bibliography management, and general Word formatting tools are outside the initial product.

## Acceptance

The first usable release is ready when a Word user can open the add-in, type a supported reference followed by a delimiter, see the correct temporary annotation without a network request, optionally preview one authorized DBS passage on hover, click or use `Alt+Down` to replace only that exact reference after its final observable freshness check, Undo back to the reference, and receive clear failure behavior for observable stale text, unavailable DBS, invalid references, unsupported hosts, and a closed/reopened task pane. The task-pane Insert control remains a keyboard-reachable alternative, and the documented WordApi 1.7 check-then-replace coauthor timing limitation remains explicit.
