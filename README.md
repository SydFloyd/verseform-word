# Verseform for Word

Verseform for Word is a small Microsoft Word add-in for people who write with Scripture. A completed reference becomes interactive after a delimiter; the user can preview authorized Digital Bible Society text and explicitly replace that reference with the passage and an editable citation.

This is the project seed. The host-capability task pane and production-shaped XML manifest are present; the first real slice is the Word on the web annotation proof described in `outputs/verseform-word-roadmap.md`.

## Why one add-in

Office.js task-pane add-ins use a web application plus a manifest and can run in Word on the web and supported desktop clients. We will prove Word on the web first, then validate the same add-in on Word for Windows rather than build a separate plugin. The add-in-only XML manifest is intentional: Microsoft's unified manifest for Word remains preview-only for production add-ins.

The central interaction targets WordApi 1.7 annotations. Word can underline an affected text range and report hover, click, and popup-action events. Those APIs require Word connected to a Microsoft 365 subscription. VFW-010 treats the exact interaction as a host feasibility gate because Word, not the add-in, owns the document canvas and native annotation popup.

## Trust boundary

- Detection is local and deterministic. It never calls DBS.
- A preview or insertion may send only the selected translation, canonical book, and chapter to the DBS ARC service.
- Document prose, file names, account identity, and surrounding text are never sent by Verseform.
- No account, analytics, telemetry, crash upload, advertising, or bundled Bible is planned.
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

Use `npm run check` for the local code and structural manifest gate. Microsoft Marketplace validation remains a release gate; do not claim the Word interaction until the VFW-010 online acceptance proof passes in the real host.

## Project authorities

Read these in order:

1. `WORK.md`
2. `outputs/verseform-word-requirements.md`
3. `outputs/verseform-word-system-design.md`
4. `outputs/verseform-word-decisions.md`
5. `outputs/verseform-word-roadmap.md`

## Stewardship

This project is being prepared with gratitude for Digital Bible Society and the people who serve others through Scripture-centered writing. Accuracy, privacy, attribution, and calm usability are part of the service, not polish added afterward.

Microsoft Word is a trademark of Microsoft Corporation. Verseform is not presented as endorsed by Microsoft. DBS relationship, repository ownership, final branding, and source license remain explicit owner decisions in the decision log.
