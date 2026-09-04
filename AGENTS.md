# Verseform for Word agent contract

## Mission

Build a small, trustworthy Word add-in that makes Scripture insertion feel native. After a delimiter, a supported reference is detected locally and becomes interactive; preview and insertion use authorized DBS text, and only an explicit fresh action may alter the document.

This tool serves people preparing Scripture-centered writing. Treat their words carefully. Precision, privacy, attribution, accessibility, and a calm interface are part of faithful stewardship.

## Read order and authority

Start every work session with `WORK.md`, then read only what the active item needs:

1. `outputs/verseform-word-requirements.md` — product behavior and scope; highest authority.
2. `outputs/verseform-word-system-design.md` — architecture and trust boundaries.
3. `outputs/verseform-word-decisions.md` — durable accepted and pending choices.
4. `outputs/verseform-word-roadmap.md` — sequence and exit proof.

When authorities conflict, follow that order and repair the lower authority in the same change.

## Operating loop

1. Inspect `git status` and `WORK.md`; preserve unrelated owner changes.
2. Select exactly one roadmap item and state its user outcome and proof before editing.
3. Trace requirement and invariant → kernel event/effect → one port → Word/task-pane view → cheapest truthful proof.
4. Implement the smallest vertical slice. Do not build host infrastructure without a user-visible consumer.
5. Put parsing, validation, freshness, and replacement planning in the pure core. Keep Office.js, DBS, storage, and clock behavior behind named ports.
6. Add the cheapest regression proof. Use pure tests and a fake Word gateway heavily; reserve real Word on the web and Windows runs for host claims.
7. Run focused checks, then `npm run check` once before completion. Run manifest validation when its truth changes.
8. Update `WORK.md`, the roadmap, and decisions only when their truth changes.

## Invariants

- Detection is local, deterministic, delimiter-triggered, and makes no network request.
- The add-in never sends document prose, file names, account identity, or surrounding text to DBS.
- Preview and insertion may request only normalized translation, book, and chapter coordinates.
- No bundled WEB or other Bible text. Offline detection is allowed; offline passage insertion is not.
- No replacement without explicit activation. A stale paragraph, moved reference, changed translation, or superseded response cannot replace text.
- Word owns document rendering, selection, undo, and annotation UI. Never emulate Word's editor in the task pane.
- Remote text is bounded, schema-checked, normalized to plain text, and never injected as provider HTML.
- Required translation attribution remains attached to inserted text and visible to the user.
- One source tree and add-in manifest should serve Word on the web and Word desktop wherever the required API set is supported.
- No telemetry, crash upload, accounts, advertising, remote configuration, or background document scanning.

## Scope

Word on the web is the first proof host. Word for Windows follows using the same Office.js application. Do not add Google Docs, Outlook, PowerPoint, browser-page injection, VBA/VSTO, bundled Bibles, collaboration features, or AI-generated content without an explicit requirements change.

The DBS webpage detector is reference material, not a mandatory dependency. Reuse only a pure, well-tested portion that fits this system and preserve its MIT notice; never load its remote script or DOM-mutating popup implementation into the Word host.

## Completion standard

A slice is complete only when its roadmap exit proof passes, failure/cancellation/stale paths are covered in proportion to risk, accessibility is preserved, real-host claims were actually observed, and `WORK.md` gives the next agent an accurate starting point. Do not publish, change DBS production state, choose a legal license, or handle production credentials without owner authority.
