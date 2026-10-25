# Changelog

## 2.0.2 - 2026-10-25
- Redesigned the GitHub README with a vector hero banner, two Mermaid diagrams, an API verification walkthrough, clearer installation steps, and explicit evidence/limitations.
- Reconciled the downloadable skill package, the repository layout, and the named release archive; retained namespace-aware API preflight from v2.0.1.
- Updated the package file-integrity manifest for the revised documentation and added SVG banner.
- Reused the exact previous offline Microsoft knowledge archive without changes; no VBA/Excel compilation or runtime testing is claimed for this documentation-focused release.

## 2.0.1 - 2026-10-25
- Audited the named VBA, Excel, MSForms, Office and VBIDE objects against the supplied snapshot; all 15 named entries have source documentation.
- Added read-only namespace-aware `api_preflight.py` and regression tests for cross-library mistakes and unconfirmed MSForms members.
- Routed agents to original Microsoft source and grouped VBIDE chapters; mitigated 8-result and PowerShell/CMD confusion.
- Preserved the Microsoft source archive and existing knowledge ZIP unchanged. No Excel compilation or runtime testing was performed.

## 2.0.0 - 2026-10-23
- Extended the existing skill, without replacing its 19 expert modules, seven
  workflows, eight original Python tool/helper files, or VBA assets.
- Imported every file from the user-supplied VBA-Docs ZIP; indexed all 38,187
  Markdown documents and retained full original text, images and licenses.
- Added offline full-text/exact API search, full-source read, line/section reads,
  source-integrity checks, local include/link resolution and curated-ID aliases.
- Added safe ZIP/directory import into a new portable knowledge folder.
- Added mandatory local source lookup before uncertain API use.
- Added real-snapshot and hostile-input tests plus OS-enforced offline test runner.
- Split installable skill and knowledge payload to respect the skill ZIP limit.
- Reported source omissions and untested Excel/live-agent behavior separately.

## 1.0.0 - 2026-10-23
- Created the modular Excel VBA Expert skill with 19 domain references, 7 workflows,
  70 reviewed source records and progressive loading.
- Added offline search, local repository inventory and read-only VBA/workbook checks.
- Added explicitly opt-in URL checking and new-blank-workbook COM smoke harness.
- Added original VBA examples, scenario contracts and Windows acceptance procedure.
- Separated source review, Python/mocked tests and unexecuted Excel/agent tests.
- Recorded incomplete upstream inventory rather than claiming exhaustive coverage.
