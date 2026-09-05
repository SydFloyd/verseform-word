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

## Word for Windows validation

Use the same `manifest.xml` and source application in desktop Word. After trusting the development certificate and closing any earlier debug session, run:

```powershell
npm run desktop:start
```

This invokes Microsoft's pinned desktop sideload helper on demand, starts the HTTPS development server, registers the add-in for Word, and opens a generated test document. The helper is deliberately not a persistent project dependency: its current development-only graph carries known high-severity audit findings and none of it ships in the add-in. Always end the session with:

```powershell
npm run desktop:stop
```

The Windows walking proof uses the same user flow as Word on the web:

1. Open **Verseform** from Word's Home ribbon or **Add-ins** menu and confirm the task pane reports **Word is ready** and loads authorized translations with NASB preferred.
2. Type `🙂 John 3:16 ` and confirm only `John 3:16` becomes a temporary annotation after the final space.
3. Activate the reference once by click and once by Word's `Alt+Down` path; confirm the exact NASB preview and Lockman attribution appear without changing the document.
4. Choose **Insert passage**, confirm passage, editable citation, and visible attribution replace only that occurrence, then use one Word Undo to restore the reference.
5. Repeat with duplicate references in one paragraph, clear the local Scripture cache, and confirm a later activation refetches rather than changing prose early.
6. Change a reference while a preview is pending and confirm Verseform refuses the stale result. Close and reopen the pane and confirm delimiter detection resumes safely.
7. Repeat keyboard insertion, cancellation, and focus recovery with Windows forced colors and a screen reader before public release.

## Production package

Production values are intentionally supplied at release time so a localhost URL, guessed DBS identity, or provisional support address cannot enter a public manifest:

```powershell
npm run build
npm run manifest:production -- -PublicBaseUrl "https://approved.example/verseform-word" -SupportUrl "https://approved.example/support" -ProviderName "Approved publisher"
npm run validate:release
```

The generated manifest is `dist/verseform-for-word-manifest.xml`. The builder requires credential-free HTTPS URLs, replaces all eight localhost app resources with one approved public base URL, preserves the reviewed WordApi 1.7/permission contract, and runs the production validator. `npm run check` also builds a non-routable fixture manifest and proves the distributable web assets retain the reviewed CSP and contain no localhost origin.

The owner and DBS must still resolve `D-011`: public HTTPS host and log retention, support URL, and Marketplace publisher identity. After those real values are known, the hosted assets and generated manifest must pass Microsoft's external manifest/Marketplace validation and the same Word web/Windows walk. A test catalog or sideload registration is not a production deployment.

## Known limits and support

- WordApi 1.7 and a connected Microsoft 365 subscription are required for temporary annotations.
- Click and `Alt+Down` are supported activation paths. Hover is best effort because Word on the web did not reliably deliver its documented event.
- Detection remains local when DBS is unavailable, but preview and insertion need a connection; no fallback Bible is bundled.
- WordApi 1.7 cannot make the final check-and-replace atomic against a coauthor edit that lands between Word synchronization boundaries. Every stale state Verseform can observe still fails closed.
- Verseform collects no diagnostics. Support is direct and voluntary; share the app version, Word host/version, reproduction steps, and non-sensitive screenshots if useful, but never private writing or Word documents.

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
