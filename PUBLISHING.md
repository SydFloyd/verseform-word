# Verseform for Word — DBS publication handoff

Researched against current Microsoft documentation on 2026-09-07. Recheck linked policies immediately before submission because Partner Center fields and certification rules change.

## Recommended release path

Publish **Verseform for Word** as a free **Office Add-in** through Partner Center's **Microsoft 365 and Copilot** program. A successful Microsoft Marketplace publication also makes the add-in discoverable in Word's in-product add-in experience. This is the clean public-distribution route; the upload package is the production XML manifest, while the add-in web application remains on DBS-controlled HTTPS hosting.

For a DBS-only or invited-user pilot before certification, use **Microsoft 365 admin center → Settings → Integrated apps → Add-ins → Deploy Add-in** and upload the production manifest to selected users or groups. Microsoft recommends Integrated Apps for organization-controlled deployment, but it is tenant distribution—not a public substitute for Marketplace. Allow up to 24 hours for a centrally deployed add-in to appear.

An owner-approved friend pilot is prepared at `https://verseform-word.kmproto.com/install`. Its downloadable manifest uses the same public HTTPS application origin and retains **Verseform Project** as the temporary provider name. Manual upload in Word on the web persists only in that browser profile; clearing its data or changing browsers requires another upload, and it does not register the add-in in Word Desktop. This is useful field testing, not production distribution or a DBS publisher claim.

Do not use a SharePoint app catalog as the primary handoff. Microsoft recommends Integrated Apps for cloud tenants, SharePoint catalogs do not support Mac, and neither route makes the add-in publicly discoverable.

Primary Microsoft references:

- [Office Add-in submission guide](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/add-in-submission-guide)
- [Microsoft 365 app publishing checklist](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
- [Marketplace certification policies, including sections 1100 and 1120](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies)
- [Open an Office account in Partner Center](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/open-a-developer-account)
- [Validate an Office Add-in manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest)
- [Deploy Office Add-ins in Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide)
- [Create effective Marketplace listings](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings)
- [Craft effective Marketplace images](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/craft-effective-appsource-store-images)

## What is already prepared

- One cross-host TypeScript/Office.js add-in with add-in-only XML manifest ID `82d488fc-c5a7-4aa5-8133-d56c32ef9a69`.
- A deterministic production-manifest builder that replaces localhost with one HTTPS application base, injects the final publisher and support URL, and rejects credentials or insecure URLs.
- Local repository gate: strict typecheck, deterministic tests, production build, manifest contract, CSP, no-localhost asset check, and production transformation.
- Microsoft-hosted Office.js `/1/hosted/office.js`, WordApi 1.7, SharedRuntime 1.1, `ReadWriteDocument`, add-in commands, 32/64-pixel listing icons plus command icon assets, and a valid support field in the manifest template.
- MIT source license, third-party notices, explicit DBS attribution, and an app-specific privacy document.
- No account, Microsoft Entra ID, SSO, advertising, analytics, telemetry, in-app purchase, paid external service, or bundled Bible.
- Tested Word on the web and Word for Windows behavior, including Fill, guarded replacement, NASB/KJV retrieval, local cache, failure paths, and one-step Undo.

## Decisions DBS must own before a production package can exist

1. **Publisher account and authority.** Identify the DBS Microsoft Entra work account and the person authorized to accept Microsoft's agreements. Enroll that publisher in Partner Center's Microsoft 365 and Copilot program and complete email, employment, and business verification. The publisher selected when the offer is created cannot later be changed.
2. **Exact publisher name.** Supply the Partner Center publisher display name. It must match the production manifest's `ProviderName`; do not submit the development value `Verseform Project`.
3. **Production host.** Choose a durable DBS-controlled HTTPS base for the static `dist` assets, with a valid public certificate, no authentication, stable URLs, suitable uptime, and no `no-cache`/`no-store` headers on icon resources. Document ordinary HTTP log access and retention in `PRIVACY.md` before launch.
4. **Public support page.** Supply a working HTTPS page—not only an email address—with installation, requirements, troubleshooting, privacy-safe contact instructions, and response ownership.
5. **Privacy URL.** Publish the final app-specific `PRIVACY.md` at a durable HTTPS URL. It must name Verseform for Word, explain personal-information handling and the coordinate-only DBS request, and reflect the chosen host/log policy.
6. **EULA.** With legal review, choose either Microsoft's Standard Contract or a DBS-hosted HTTPS EULA. Partner Center warns that choosing the Standard Contract cannot be reversed after publication.
7. **Brand approval.** Confirm the final product name, DBS credit/relationship wording, app icon, Marketplace logo, and who owns future updates. Reserve **Verseform for Word** in Partner Center before assuming it is available.
8. **Release timing and availability.** Choose regions and the first availability date carefully; Microsoft's submission guide says the schedule choice cannot be changed after first publish.

## Material blockers before Marketplace submission

- **Mac host proof:** Microsoft policy 1120.3 requires the add-in to work on every platform admitted by its host/requirement-set manifest. This manifest currently admits Word on Windows, Word on the web, and Word on Mac. Microsoft doesn't allow a Marketplace add-in to use a runtime platform check merely to reject Mac. Obtain current Word for Mac and Safari/Word-web passes for the complete certification walk, or do not submit publicly yet.
- **Windows accessibility proof:** finish forced-colors, touch-only operation, keyboard access through Word's native Fill command, and a screen-reader walk. Policy 1120.3 explicitly requires touch-only operation.
- **Selected-text context menu proof:** confirm exact selection, two-reference ambiguity refusal, and Undo in release hosts.
- **Production-host proof:** deploy the exact built assets, generate the manifest from final values, run production validation, and repeat the web/Windows/Mac test matrix from that public origin.
- **Marketplace media:** create an approved 216–350-pixel square PNG logo. Microsoft's current general image policy calls for 1280×720 PNG screenshots, while its Office listing image guide asks for at least one 1366×768 image no larger than 1,024 KB. Export the same approved scene at both sizes, keep text legible and personal data absent, and upload the size accepted by the live Partner Center field. One screenshot is required for an Office Add-in; prepare three focused scenes to give DBS a stronger listing.
- **Legal/support publication:** final privacy, support, and EULA URLs must return successful HTTPS pages before submission.

## Produce the final package

From a clean checkout with Node.js 22:

```powershell
npm ci
npm run check
npm run build
npm run manifest:production -- -PublicBaseUrl "https://DBS-APP-BASE" -SupportUrl "https://DBS-SUPPORT-PAGE" -ProviderName "EXACT PARTNER CENTER PUBLISHER"
npm run validate:marketplace
```

Deploy the complete contents of `dist` at `PublicBaseUrl` before running the production validator because Microsoft must reach the manifest's source, command, icon, and support URLs. Upload `dist/verseform-for-word-manifest.xml` on the Partner Center **Packages** page. Do not upload the localhost `manifest.xml`, the source repository, or a desktop installer.

Before tagging the release, verify that:

- `git status --short` is empty and the tested commit is recorded.
- The production manifest has the same stable app GUID, incremented version, exact publisher name, HTTPS support URL, and no localhost references.
- Public `index.html`, hashed assets, 32/64/128 icons, privacy, support, and EULA pages return HTTP 200 without authentication.
- Icon responses permit caching and no response redirects to login, an error page, or a different unreviewed origin.
- The final CSP allows only the reviewed self, Office.js, and DBS ARC boundaries.
- The release is tested from the hosted origin—not Vite or a developer certificate—in Word web, Windows, and Mac.

## Partner Center form plan

1. Sign in with the verified DBS work account, open **Marketplace offers**, choose **Microsoft 365 and Copilot**, then **New offer → Office Add-in**.
2. Check availability for **Verseform for Word**, choose the final DBS publisher, and create the offer. The publisher association is permanent.
3. Product setup: no Apple Store listing unless DBS separately elects and tests it; no Entra ID/SSO; no additional purchases; no lead-management CRM.
4. Packages: upload only the production XML manifest and resolve every automated check.
5. Properties: select one to three accurate categories and at most two genuine industries. Choose the reviewed EULA path, then enter the public privacy and support HTTPS URLs.
6. Marketplace listings: English (`en-US`) only for this release. Add the name, summary, description, keywords, logo, at least one screenshot, and captions. Do not add untranslated locales or multilingual detection claims.
7. Availability: select the approved regions and launch date.
8. Paste the certification notes below and optionally upload a reviewer PDF containing the same walk with screenshots. Reviewers must not need to contact a developer or receive time-sensitive credentials.
9. Select **Review and publish**, address automated/certification findings, inspect the certified preview, then have DBS explicitly approve **Go live**. Plan four to six weeks and at least one resubmission; Microsoft says initial issue feedback commonly arrives within three to four business days.

## Draft Marketplace listing (`en-US`)

**Name:** Verseform for Word

**Summary:** Insert authorized Scripture in Word without interrupting your writing.

**Suggested keywords:** Bible, Scripture, verse, sermon, ministry, teaching, Christian writing, Digital Bible Society

**Description:**

Verseform for Word helps pastors, teachers, ministry teams, and Scripture-centered writers stay focused inside Microsoft Word. Type a supported English Bible reference, leave the cursor in or immediately after it, and choose **Fill Scripture**. Verseform retrieves the authorized passage from Digital Bible Society and replaces only that reference with the passage and a compact, editable translation citation. One ordinary Word Undo restores the reference.

The add-in supports individual verses and complete same-chapter ranges, including Psalm 119:1–176. Choose from translations authorized by the Digital Bible Society catalog; NASB is preferred when available, or Verseform remembers the authorized translation you select. **Preview & Settings** provides an optional passage preview, translation details, the provider's complete notice, cache controls, privacy information, and concise help. The pane is not required for the primary Fill Scripture command.

Privacy is part of the design. Reference detection happens locally inside the add-in and never contacts Digital Bible Society. Only an explicit preview or insertion may request the selected translation identifier, canonical book identifier, and chapter number. Verseform does not send document prose, document names, Microsoft identity, or surrounding text. It has no Verseform account, advertising, analytics, telemetry, crash upload, or bundled Bible. Authorized catalog and chapter responses may be retained in a small, bounded local cache that the user can clear.

Verseform for Word is free and requires no additional purchase or external account. Passage preview and insertion require an internet connection to the Digital Bible Society service. Temporary inline annotations require a connected Microsoft 365 subscription and a supported Word host; **Fill Scripture** remains the explicit insertion action. The initial release recognizes English Bible book names and references. Translation-specific copyright and document-level attribution requirements remain visible in Preview & Settings.

Scripture service is provided by Digital Bible Society. Microsoft and translation publishers do not endorse Verseform unless separately stated in writing.

## Draft notes for Microsoft certification

Verseform for Word is a free Word task-pane/command add-in. It has no sign-in, SSO, account, license key, purchase, subscription, or test credentials. It uses the public Digital Bible Society ARC service at `https://arc.dbs.org/api/bible-text/`; reviewers do not need a DBS account. No Bible is bundled.

Test with a connected Microsoft 365 subscription and a Word host that supports WordApi 1.7 and SharedRuntime 1.1:

1. Open a blank document. On Home → Scripture, confirm **Fill Scripture** and **Preview & Settings** are available and the pane is not required.
2. Type `John 3:16`, leave the cursor immediately after the reference, and choose **Fill Scripture**. Confirm only that reference becomes the selected authorized passage plus an editable citation.
3. Use one Word Undo and confirm `John 3:16` returns.
4. Open **Preview & Settings**. Confirm NASB is selected when authorized, the help/privacy text and full translation notice are visible, and another authorized translation can be selected.
5. Type `Psalm 119:1-176` and use Fill. Confirm the complete chapter is accepted. Undo it. Type `Psalm 119:1-177`; confirm Verseform refuses it without changing the document.
6. Select one exact reference and invoke **Fill Scripture** from Word's text context menu. Confirm a selection touching two references is refused without a change.
7. Start an uncached preview/insertion and edit the reference before it completes. Confirm Verseform refuses the stale result and does not replace prose.
8. Disconnect network access, use Fill on an uncached valid reference, and confirm the optional pane explains that Scripture text needs a connection while the document remains unchanged. Restore network access before continuing.
9. In Preview & Settings, clear the local Scripture cache and confirm the selected translation preference remains.
10. Repeat the core Fill/Undo, task-pane, keyboard, touch-only, forced-color, and screen-reader walk on Word for Windows, Word on the web in current Edge/Chrome/Firefox/Safari as applicable, and Word for Mac.

Data disclosure: local detection reads only the relevant Word paragraph. An uncached explicit preview or insertion sends only translation ID, canonical book ID, and chapter number to DBS. Document prose, file name, title, identity, surrounding text, and request history are not sent. The app requests `ReadWriteDocument` to read the target paragraph, create temporary annotations, and perform the user-requested replacement. The privacy policy contains the complete cache and hosting disclosure.

## Handoff inventory

Deliver to DBS:

- Repository URL and exact release commit/tag.
- `dist` static web assets and `dist/verseform-for-word-manifest.xml` generated from DBS's final values.
- Public application, support, privacy, and EULA URLs.
- Partner Center publisher/Seller ID and offer alias recorded outside the repository.
- Approved listing copy, logo, screenshots/captions, categories, regions, and availability date.
- Completed hosted-origin test matrix with Word/web/OS/browser versions and tester/date.
- Microsoft validator output, Partner Center automated report, certification correspondence, and final public Marketplace URL.
- Named owners for hosting, support, privacy review, DBS API coordination, Partner Center notifications, and future manifest/version updates.
