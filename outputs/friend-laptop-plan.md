# Verseform: one-person Word installation plan

Updated 2 November 2026. Goal: a reliable, always-available tool for one friend, not a public release. Existing source: `/home/admin/Projects/verseform-word`, commit `f743492` (`record live friend pilot`). DBS permission for NASB is already confirmed by the owner; reuse the existing authorized DBS integration and its attribution behavior.

## What already exists

The existing add-in detects references locally, marks them with Word annotations, uses DBS ARC for Scripture, defaults to NASB unless a preference is saved, and inserts a passage with a citation after explicit activation. It includes a version selector, bounded local cache and stale-document checks. It currently **replaces** the reference with Scripture and citation; confirm that this is the desired behavior. Its live friend-pilot hosting is documented at https://verseform-word.kmproto.com/install; that page describes Word on the web, not a desktop installation. Confirm the host remains available and who maintains it before relying on it.

The manifest requires WordApi 1.7 and SharedRuntime 1.1; annotations also require a connected Microsoft 365 subscription. It cannot simply be installed in a Word that lacks these capabilities. A previous Windows test on Word 16.0.20326.20132 proved ribbon **Fill Scripture** and Undo with the pane closed; annotation clicks did not reach the add-in while the pane was hidden. Clicking the annotation worked with the pane visible. This is evidence for that host, not yet this friend's laptop. The entry point currently turns document startup behavior off; Fill starts the runtime on demand. Thus the existing implementation is not yet proved always active on a fresh document or after a Word restart.

## Before the visit

1. Ask for a photo or transcription of Word's **File > Account > About Word**: product name, version/build, license and 32/64-bit; record OS and whether it is Windows or Mac. The 16.x number alone does not mean the perpetual 2016 product. Confirm whether he has a connected Microsoft 365 subscription, internet access, and any device or account restrictions.
2. Confirm NASB edition and insertion preference; obtain representative references and a sample disposable document. Verify the existing DBS permission covers this specific one-person use and edition under the arrangement already held, without starting a new licensing exercise.
3. Take a pinned checkout, the installation instructions, and a copy of the current pilot manifest. Review current hosting ownership, deployment method, HTTPS renewal, connectivity and the ability to restore a known working version. A hosted web add-in still depends on that host and DBS; the owner should know who maintains them. Do not require a developer server to be running daily.

## At the laptop: compatibility and startup first

1. Verify host requirements using a small diagnostic or Word's actual behavior, including `isSetSupported('WordApi', '1.7')` and `isSetSupported('SharedRuntime', '1.1')`. Confirm WebView and HTTPS/DBS access. Do not change the existing manifest's minimum requirements and assume the code will work.
2. If compatible, install the existing hosted manifest using a desktop method supported by his account and machine. The existing web sideload is not a desktop installation. Microsoft's Windows network-share catalog is a test sideload method; confirm persistence, trust, restart behavior and recovery on this laptop before choosing it. Do not use Marketplace publication as a prerequisite.
3. Test how to make detection active: shared runtime lasts after pane close but closes with the document. Microsoft documents `setStartupBehavior(load)` for the **current document**, not a global switch for every new or existing document. The current code explicitly sets startup to `none`. Prototype the smallest appropriate change and test new blank documents, existing files, close/reopen and full Word/laptop restart. Investigate a suitable template or other supported per-machine activation only if needed; verify behavior rather than promise automatic startup across all documents. If that cannot be achieved on the host, present the actual startup action and assess an alternative architecture with the owner.
4. With the pane closed, type `John 3:16` and test the recognition mark, clicking it, and **Fill Scripture**. Then test with the pane open. Retain Fill as the proven fallback if hidden-pane click events do not arrive. Never claim click-to-insert works on this machine until observed.
5. If Word lacks the required APIs, decide with the owner whether a supported Word upgrade is acceptable, or investigate a separate Windows-specific approach if the machine is Windows. A scan-and-insert pane may be a simpler fallback but will not itself satisfy always-active inline detection. Do not spend the visit silently trying to make the current annotation implementation work on an unsupported host.

## Finish and handoff

- Test NASB text and edition, version switching/persistence, single and range references, duplicate and invalid references, citation, undo, and saving/reopening. Change a reference while a lookup is pending; verify no incorrect replacement. Test offline/DBS failure and recovery.
- Verify first use after a cold laptop restart, a fresh document, an existing document and a closed pane. Record whether startup is automatic in each case, and what action recovers it.
- Give the friend a short usage and troubleshooting note: installation, how to insert, translation setting, internet dependency, how to restore the add-in and a contact for hosting/maintenance. Keep the last known working manifest and build and record the installed version. Avoid unrelated public-release and Marketplace work.

## References

- [Microsoft: run code when a document opens](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/run-code-on-document-open)
- [Microsoft: Windows network-share sideload for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [Microsoft: Office API runtime checks](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-api-requirements-runtime)
- [Microsoft: Office 2016 end of support](https://learn.microsoft.com/en-us/deployoffice/endofsupport/resources)
