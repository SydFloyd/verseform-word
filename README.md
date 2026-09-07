# Verseform for Word

Verseform for Word is a small Microsoft Word add-in for people who write with Scripture. Type a reference, leave the cursor in or immediately after it, and choose **Fill Scripture** from Word's Home ribbon; the exact occurrence becomes authorized Digital Bible Society text with a compact editable citation after a final freshness check. The preview/settings pane is optional.

The project seed, Word interaction proof, and authorized DBS walking slice are complete. The add-in detects references locally, creates exact temporary annotations, loads the authorized translation catalog, and replaces only an explicitly activated, observably fresh occurrence with passage and editable translation citation. NASB is preferred unless the user has saved another authorized translation. An optional pane provides preview, settings, and the selected translation's complete provider notice.

## Why one add-in

Office.js add-ins use a web application plus a manifest and can run in Word on the web and supported desktop clients. Verseform uses one long SharedRuntime 1.1 runtime so ribbon commands and the optional pane share state and detection continues after the pane closes. We proved Word on the web first and validate the same add-in on Word for Windows rather than build a separate plugin. The add-in-only XML manifest is intentional: Microsoft's unified manifest for Word remains preview-only for production add-ins.

The central interaction uses WordApi 1.7 annotations for inline feedback and one explicit **Fill Scripture** command for reliable pane-free insertion. Fill is available on the ribbon and selected-text context menu, targets one reference at or immediately before the cursor, and works without a prior annotation event or open pane. Annotation click and `Alt+Down` remain conveniences where Word delivers them; hover is optional preview-only behavior. The task-pane Insert control remains an accessible alternative. WordApi 1.8 critique suggestions are not repurposed because Word owns their document replacement. Observable stale, unknown, changed, and ambiguous states fail closed; WordApi 1.7's narrow non-atomic coauthor timing limitation remains documented.

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
npm run build
npm run desktop:serve
```

Then open Word on the web, choose **Add-ins → Advanced → Upload My Add-in**, and upload `manifest.xml`. The local server must remain running at `https://localhost:3000`.

Use `npm run check` for the local code and structural manifest gate. VFW-010 and VFW-020 passed their Word on the web host proofs, including live NASB/KJV retrieval, local caching and clearing, editable attribution insertion, one-step Undo, stale-response rejection, and provider-unavailable behavior without prose loss. VFW-030 now adds a versioned English corpus covering every approved canonical name/alias, strict and fuzzy behavior, false positives, mixed-direction UTF-16 offsets, provider isolation, and a 100,000-unit performance budget. Multilingual detection is deliberately deferred pending DBS review and English-pilot feedback. Production hosting/Marketplace validation, accessibility validation, and the same-source Word for Windows walk remain release gates.

## Word for Windows validation

Use the same `manifest.xml` and source application in desktop Word. After trusting the development certificate and closing any earlier debug session, run:

```powershell
npm run desktop:start
```

This builds the production assets, invokes Microsoft's pinned desktop sideload helper on demand, serves the built assets over local HTTPS, registers the add-in for Word, and opens a generated test document. It deliberately avoids Vite's development client: the production CSP blocks that client's injected styles and live-reload WebSocket. After source changes, rebuild and reload the add-in; do not claim a hot reload. The helper is deliberately not a persistent project dependency: its current development-only graph carries known high-severity audit findings and none of it ships in the add-in. Save and close the test document before ending the session with:

```powershell
npm run desktop:stop
```

The following is the Windows walking proof. Ribbon Fill and one-step Undo have passed on the installed host; the remaining context-menu and accessibility steps are release gates:

1. Choose **Enable Verseform** from Word's Home ribbon or **Add-ins** menu. Confirm no task pane is required.
2. With the pane closed, type `John 3:16` and leave the cursor at its end. Choose **Fill Scripture** and confirm passage plus the compact editable NASB citation replace only that occurrence even without a trailing delimiter.
3. Use one Word Undo to restore the reference. Repeat by selecting exactly one reference and choosing **Fill Scripture** from Word's text context menu. Confirm a selection touching two references is refused without a document change.
4. Type `🙂 James 4:17`, press Enter, and confirm the preceding reference becomes annotated at the exact UTF-16 offset. Open **Preview & settings** only as an option, confirm hover can preview when delivered, the full Lockman notice is visible, and closing the pane does not stop later detection or ribbon Fill.
5. Repeat with duplicate references in one paragraph, clear the local Scripture cache, and confirm a later activation refetches rather than changing prose early.
6. Change a reference while a preview is pending and confirm Verseform refuses the stale result. Close and reopen the pane and confirm delimiter detection resumes safely.
7. Repeat ribbon/context-menu insertion, cancellation, and focus recovery with Windows forced colors and a screen reader before public release. Do not claim a custom simultaneous shortcut until Word registers it in the release host.

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
- On Word for Windows 16.0.20326.20132 / WebView2 152.0.4191.66, the shared runtime and paragraph events survive pane close, but native annotation click/`Alt+Down` events do not arrive while hidden. **Fill Scripture** contains that host boundary and is the required pane-free path; the same annotation conveniences remain available where Word delivers them.
- Ribbon Fill is one explicit action after typing. It locally resolves the reference at or immediately before the cursor, requests only canonical DBS coordinates, and performs the same guarded single replacement. Hover is passive while hidden and optional preview while the pane is open.
- No custom add-in shortcut is advertised. Three clean sideload reloads—including Microsoft's current sample key and JSON shape—did not register one in this installed Word host; a future shortcut requires release-host proof rather than silently overriding Office preferences.
- SharedRuntime 1.1 is required so detection continues while the optional pane is closed. **Enable Verseform** opts the current document into pane-free startup on its next open.
- Insertions include a compact translation citation, not a repeated full copyright paragraph. The optional pane exposes the selected translation's complete provider notice for any document-level attribution the writer needs.
- Detection remains local when DBS is unavailable, but preview and insertion need a connection; no fallback Bible is bundled. An explicit insertion failure opens the optional pane so the exact no-change error is visible.
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
