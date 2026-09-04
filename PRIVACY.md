# Verseform for Word privacy

Verseform for Word is designed without an account, advertising, analytics, telemetry, or crash upload.

## Document access

The add-in requests Word's `ReadWriteDocument` permission because it must read changed paragraph text, add temporary reference annotations, and replace a reference only after the user explicitly chooses insertion. It does not upload the document, its title, surrounding prose, Microsoft identity, or coauthor information.

Detection runs locally inside the add-in. Word annotations are owned by the Word host, are not persisted by the annotation API, and rely on a Microsoft 365 service. Microsoft 365 and Word on the web have their own data handling outside Verseform's control.

## DBS requests

A catalog request may load the translations authorized by Digital Bible Society. An uncached preview or insertion may request only the selected translation identifier, canonical book identifier, and chapter number from the DBS ARC API. Provider responses are treated as untrusted data and never executed as HTML.

No Bible is bundled. Without DBS connectivity, reference detection may continue but passage preview and insertion are unavailable.

## Hosting and diagnostics

The production add-in will be served over HTTPS. The host may necessarily process ordinary web request metadata such as IP address and user agent; production host, retention, and access policy must be settled before public release. Verseform will not add product analytics or document-content logging.

User feedback is direct and voluntary. Never ask users to submit private writing or Word documents as diagnostics.
