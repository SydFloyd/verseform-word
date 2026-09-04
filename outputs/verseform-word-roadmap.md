# Verseform for Word — Tight Roadmap

Each item is one user outcome with behavior, failure handling, and proof. Status is `done`, `next`, `queued`, or `blocked`.

| ID | Status | Outcome | Exit proof |
|---|---|---|---|
| `VFW-000` | done | Seed one coherent project: name it, define product/trust contracts, choose the cross-host Office.js shape, add a production-shaped XML manifest and capability task pane, and record the Word annotation feasibility gate and owner decisions. | Authorities agree; the scaffold builds and tests on Node 22; the manifest validates; no Bible, credentials, telemetry, or unproved Word behavior is present. |
| `VFW-010` | next | Prove the Word on the web walking interaction using native annotations. Detect one English reference locally after a delimiter, annotate the exact range, map hover/activation to that reference, show a fake passage preview, explicitly replace only a fresh exact occurrence with passage and citation, and Undo to the reference. | Pure/fake-host tests prove delimiter and stale rules without network; a real Microsoft 365 Word on the web run proves paragraph events, annotation offsets, hover/click identity, duplicate-reference behavior, one-transaction replacement, Undo, close/reopen lifecycle, and accessible task-pane fallback. The system design records the supported native interaction instead of assuming it. |
| `VFW-020` | queued | Connect authorized DBS translations without adding document-data authority. Load the bounded catalog, prefer saved/NASB selection, fetch and normalize passages on demand, insert editable citation/attribution, and fail clearly offline. | Provider contracts and recorded payloads prove bounds and text normalization; a browser network assertion exposes only allowlisted catalog/chapter coordinates; real Word proves preview, insertion, stale response rejection, translation changes, and provider failure without losing prose. |
| `VFW-030` | queued | Harden reference intelligence and decide multilingual breadth. Complete the corpus for aliases, fuzziness, ranges, invalid bounds, punctuation, URLs, duplicate occurrences, coauthor changes, right-to-left/Unicode offsets where accepted, and long paragraphs. | Versioned language corpora and performance limits pass; any DBS detector code/data reuse is locally pinned, MIT-credited, and verified at the pure boundary; detection still makes zero provider calls. |
| `VFW-040` | queued | Prepare free DBS distribution and validate the same add-in on Word for Windows. Finalize branding/license/privacy/hosting, production manifest and CSP, accessibility, Marketplace or approved catalog packaging, and support documentation. | Production HTTPS assets and manifest validate; Word on the web and Windows pass the same walking flow; Microsoft 365/WordApi limits, DBS attribution, host logging, permissions, support, and known limits are explicit; no separate desktop implementation exists. |

## Scope gates

- VFW-010 is a host feasibility gate. If stable Word APIs cannot identify and replace an exact annotated occurrence, change the interaction before provider breadth.
- VFW-020 cannot persist Scripture until `D-010` is resolved.
- VFW-030 cannot claim a language without a language-specific corpus and Word offset proof.
- VFW-040 cannot publish under DBS or choose a license/publisher identity until `D-008` and `D-011` are resolved.
