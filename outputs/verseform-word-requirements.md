# Verseform for Word — Product Requirements

## Purpose

Verseform for Word helps a person remain inside Microsoft Word while inserting accurate, authorized Scripture. The add-in should feel like a small native writing aid, not a second editor or Bible study suite.

## Central interaction

1. While the Verseform add-in is active, Word reports a changed paragraph.
2. After the user completes a supported reference and types a delimiter, Verseform scans only the relevant paragraph locally.
3. A valid reference receives a temporary Word-native annotation. Invalid coordinates may receive restrained, explanatory feedback; ordinary prose stays untouched.
4. Hovering or activating the annotation requests no text unless a preview or insertion is actually needed. VFW-010 uses the task pane for its on-demand preview and explicit Insert action. It does not assume a native popup action: the installed Office.js contract places those APIs in WordApi 1.8, above this slice's WordApi 1.7 baseline.
5. An explicit insertion action replaces exactly the activated reference with passage text followed by an editable citation.
6. Before replacement, Verseform proves that the paragraph, source range, reference, and translation still match the request. Otherwise it does nothing and explains that the writing changed.
7. The replacement is one Word transaction so Word's ordinary Undo restores the reference.

## Reference intelligence

- The initial language is English. Full names, approved abbreviations, conservative fuzzy book-name matching, single verses, and same-chapter ranges follow Verseform's proven behavior.
- Chapter and verse bounds are strict. Only book names may be fuzzy.
- Whitespace, punctuation, paragraph breaks, and common closing characters count as delimiters. An unfinished reference is not annotated.
- URLs, emails, existing generated citations, and representative prose false positives are excluded.
- Detection never contacts Word services beyond reading the changed paragraph and never contacts DBS.
- Multilingual detection is a planned capability, not an assumption. DBS's ten-language detector is candidate reference material; adopting its data or code requires corpus-backed ambiguity and bounds proof.

## Scripture and translations

- Verseform lists translations authorized by the DBS ARC API and remembers a local preference.
- If there is no preference, NASB is preferred when the catalog contains it; otherwise the first clearly available authorized translation is used.
- No WEB or other Bible is bundled. When DBS is unavailable, detection can continue but preview and insertion report that Scripture text needs a connection.
- Provider requests contain only translation ID, canonical book ID, and chapter. Document prose, file name, title, identity, and surrounding words are never sent.
- Responses are time-, size-, and schema-bounded, normalized as untrusted plain text, and never executed or rendered as provider HTML.
- Authorized catalog and Scripture responses may be cached locally. Persistent caches must be schema-versioned, bounded by age, bytes, and entries, clearable by the user, and must never contain document prose or identity.
- Inserted text includes the normalized passage and an editable citation with translation abbreviation. Required provider copyright/attribution remains visible in the task pane and travels with generated content where Word allows metadata.

## Word behavior

- One Office.js task-pane add-in targets Word on the web and supported Word desktop clients. The same core and manifest contract must be used across hosts.
- Native Word APIs own annotations, selection, replacement, formatting inheritance, Undo, coauthor behavior, and accessibility semantics.
- Temporary annotations must not persist as document content or alter saved prose.
- The add-in must behave safely when its task pane closes, event handlers are destroyed, coauthors change text, the selection moves, or the host lacks WordApi 1.7.
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

The first usable online release is ready when a Word on the web user can open the add-in, type a supported reference followed by a delimiter, see the correct temporary annotation without a network request, preview one authorized DBS passage on demand, explicitly replace only that fresh reference with passage and citation, Undo back to the reference, and receive clear failure behavior for stale text, unavailable DBS, invalid references, unsupported hosts, and a closed/reopened task pane.
