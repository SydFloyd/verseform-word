# Verseform for Word work state

Updated: 2026-09-04

## Now

- Stage: `VFW-000` project seed is complete and published.
- Active item: `VFW-010` is next — prove the real Word on the web annotation interaction before expanding the provider or multilingual surface.
- Product name: **Verseform for Word**; repository name: `verseform-word`.
- Public repository: `https://github.com/SydFloyd/verseform-word`.
- Host baseline: one Office.js task-pane application and add-in-only XML manifest target Word on the web first and Word desktop later. WordApi 1.7 and a connected Microsoft 365 subscription are required for native annotations.
- Trust baseline: detection is local; no Bible is bundled; preview/insertion may send only normalized coordinates to DBS; no document prose, telemetry, account, or crash upload leaves through Verseform.

## Last verified

- Microsoft documents Word task-pane add-ins as cross-platform across web and desktop hosts.
- Microsoft documents `onParagraphChanged` in WordApi 1.6 and native annotation insertion, hover, click, and popup-action events in WordApi 1.7. Annotation state is ephemeral and requires a connected Microsoft 365 subscription.
- Microsoft's unified manifest remains preview-only for Word production add-ins, so this seed uses an add-in-only XML manifest.
- DBS's BrowserBible verse detector is MIT-licensed and offers a pure multilingual detector, but its broader DOM/popup system and permissive reference grammar do not automatically fit the Word correctness boundary. No DBS detector source is incorporated in this seed.
- `npm run check` passes: strict typecheck, four detector/freshness tests, production build, and the local structural manifest contract.
- `npm audit --audit-level=high` reports zero vulnerabilities.

## Owner inputs before public release

- Confirm long-term GitHub ownership and whether DBS will receive or mirror the repository.
- Confirm source license; MIT is recommended for a free DBS-distributed add-in.
- Confirm whether the first usable release is English-only or must include DBS's ten detector languages.
- Confirm that DBS cache permission applies to this add-in and choose production HTTPS hosting/log-retention ownership.

## Handoff rule

Replace **Now** and **Last verified** as evidence changes. Keep detailed history in version control, not this file.
