# Verseform for Word

Verseform for Word is a small Microsoft Word add-in for people who write with Scripture. A completed reference becomes interactive after a delimiter; the user can preview authorized Digital Bible Society text and explicitly replace that reference with the passage and an editable citation.

The project seed, Word interaction proof, and authorized DBS walking slice are complete. The add-in detects references locally, creates exact temporary annotations, loads the authorized translation catalog, previews a selected passage only on activation, and replaces only an explicitly activated, observably fresh occurrence with passage, editable citation, and provider attribution. NASB is preferred unless the user has saved another authorized translation.

## Why one add-in

Office.js task-pane add-ins use a web application plus a manifest and can run in Word on the web and supported desktop clients. We will prove Word on the web first, then validate the same add-in on Word for Windows rather than build a separate plugin. The add-in-only XML manifest is intentional: Microsoft's unified manifest for Word remains preview-only for production add-ins.

The central interaction targets WordApi 1.7 annotations. Word can underline an affected text range and report click, keyboard, and hover activation; click and `Alt+Down` are the proved paths, while hover remains best effort because Word on the web did not reliably deliver it. Annotation popup-action APIs are WordApi 1.8, so this 1.7-baseline add-in keeps preview and Insert in its accessible task pane. These APIs require Word connected to a Microsoft 365 subscription. VFW-010 closed its feasibility gate with explicit safe refusal for observable stale, unknown, and post-close ambiguous states; WordApi 1.7's narrow non-atomic coauthor timing limitation remains documented.

## Trust boundary

- Detection is local and deterministic. It never calls DBS.
- A preview or insertion may send only the selected translation, canonical book, and chapter to the DBS ARC service.
- Document prose, file names, account identity, and surrounding text are never sent by Verseform.
- No account, analytics, telemetry, crash upload, advertising, or bundled Bible is planned.
- Authorized DBS catalog and Scripture responses may be kept in a bounded, clearable local cache; document prose never enters it.
- Remote content is untrusted plain data. It is bounded, validated, normalized, and never rendered as provider HTML.
- A passage replaces text only after an explicit user action and a fresh-source check.

See `PRIVACY.md` for the working privacy contract.

## Local development

Prerequisites: Node.js 22, a Microsoft 365 subscription, and a Word host that supports WordApi 1.7.

```powershell
npm install
npx office-addin-dev-certs install
npm run dev
```

Then open Word on the web, choose **Add-ins → Advanced → Upload My Add-in**, and upload `manifest.xml`. The local server must remain running at `https://localhost:3000`.

Use `npm run check` for the local code and structural manifest gate. VFW-010 and VFW-020 passed their Word on the web host proofs, including live NASB/KJV retrieval, local caching and clearing, editable attribution insertion, one-step Undo, stale-response rejection, and provider-unavailable behavior without prose loss. VFW-030 now adds a versioned English corpus covering every approved canonical name/alias, strict and fuzzy behavior, false positives, mixed-direction UTF-16 offsets, provider isolation, and a 100,000-unit performance budget. Multilingual detection is deliberately deferred pending DBS review and English-pilot feedback. Production hosting/Marketplace validation, accessibility validation, and the same-source Word for Windows walk remain release gates.

## Project authorities

Read these in order:

1. `WORK.md`
2. `outputs/verseform-word-requirements.md`
3. `outputs/verseform-word-system-design.md`
4. `outputs/verseform-word-decisions.md`
5. `outputs/verseform-word-roadmap.md`

## Stewardship

This MIT-licensed project is being prepared under `SydFloyd` with gratitude for Digital Bible Society and the people who serve others through Scripture-centered writing. Accuracy, privacy, attribution, and calm usability are part of the service, not polish added afterward. The first usable release is deliberately English-first; multilingual detection follows corpus and Word offset proof.

Microsoft Word is a trademark of Microsoft Corporation. Verseform is not presented as endorsed by Microsoft. A future DBS mirror or repository transfer, final branding, production hosting, and publisher identity require explicit coordination.
