# Verseform for Word — Decisions

Record durable choices here. Supersede a row rather than rewriting history after code depends on it.

| ID | Status | Decision | Reason |
|---|---|---|---|
| `D-001` | Accepted | Use **Verseform for Word** as the working product name and `verseform-word` as the repository name. | It preserves the trusted Verseform interaction while clearly naming the host; branding can transfer with the repository later. |
| `D-002` | Accepted | Build one TypeScript Office.js task-pane add-in for Word on the web first and supported Word desktop hosts second. Use an add-in-only XML manifest until Microsoft's unified Word manifest is production-supported. | Microsoft documents task-pane add-ins as cross-platform, while the unified Word manifest remains preview-only. Two implementations would create needless drift. |
| `D-003` | Accepted as a gated direction | Target WordApi 1.7 native annotations for inline feedback and events, but make VFW-010 prove exact offsets, hover/click behavior, duplicate references, replacement, Undo, and lifecycle in real Word on the web. | Annotations are the closest native mapping to Verseform, but Word owns the canvas and the APIs require a connected Microsoft 365 subscription. |
| `D-004` | Accepted | Own a pure local detector behind a Word-neutral contract. Start from Verseform's strict English behavior; treat DBS's MIT detector as optional reference material and never load its remote DOM/popup bundle in production. | The DBS detector brings excellent multilingual data but its webpage mutation, permissive grammar, and HTML popup are not the Word host or correctness boundary. |
| `D-005` | Accepted | Bundle no Bible text. Detection remains available offline, while preview and insertion require authorized DBS access and fail explicitly when unavailable. | This matches the owner's direction and prevents text from being mislabeled or silently substituted. |
| `D-006` | Accepted | Add no account, telemetry, crash upload, advertising, AI prose service, document upload, or background scanning. | A writing aid should not turn private Word content into operational data. |
| `D-007` | Accepted | Honor a saved authorized translation; otherwise prefer a catalog entry identified as NASB, then the first valid catalog entry. | This carries the proven Verseform default without bundling a fallback translation. |
| `D-008` | Pending owner/DBS confirmation | Select repository stewardship and source license; MIT is recommended. | Public free use does not itself grant an open-source license or establish the final legal publisher. |
| `D-009` | Pending owner/DBS confirmation | Decide whether the first usable release is English-only or includes DBS's ten current detector languages. | Multilingual names are valuable, but ambiguity, canon, directionality, and Word offsets require language-specific proof. |
| `D-010` | Pending DBS confirmation | Confirm persistent catalog/chapter caching terms for this add-in. Until then, allow only bounded in-memory deduplication per task-pane session. | Prior Verseform cache permission should not be silently broadened to a new product. |
| `D-011` | Pending owner/DBS input | Choose the production HTTPS host, server-log retention, support URL, and Microsoft Marketplace publisher identity. | These values are user-visible trust and privacy boundaries, not deployment trivia. |
