# Changelog

## 2.1.0 - 2026-10-27

### Added
- `scripts/compare_vba_source.py`: offline, read-only extraction and source comparison for VBA components in macro-enabled Office packages. Reports its exact evidence scope rather than treating binary differences as source differences.
- `scripts/compare_workbook_values.py`: offline comparison of cached cell values, preserving sheet and cell coordinates and distinguishing changed values from relocated cells.
- `references/agent-conduct.md`: maintaining a reference baseline, safe resumable work, scope-bound delegation, and accurate evidence reporting.
- Ten additional test scenario contracts, S11-S20 (20 total); additional unit tests for the comparison tools, including negative and damaged-input cases.

### Improved
- Strengthened testing guidance around non-vacuous assertions, mutation/negative tests, incomplete reads, test prerequisites, and risk-proportionate verification.
- Expanded guidance on VBA array bounds, empty arrays, qualified `Application.Run`, dependency-order aggregation, cycles and rule precedence.
- Clarified the difference between byte-level file identity, cached worksheet value equality and extracted VBA source equality.
- Refined safe deployment, source encoding, data snapshots, restoration verification, Excel COM ownership and security-preserving offline investigation.
- Documentation and release metadata reconciled with the existing v2.0.2 public presentation. Existing Microsoft knowledge archive unchanged.

### Validation and limits
- Structural validation and integrity manifest: PASS; 20 modules, 7 workflows, 14 scripts, 20 scenario contracts.
- Maintainer Windows report: 156 Python unit tests passed (1 skipped), including 31 comparison-tool tests and independent probe cases.
- Source extraction cross-checked against an independent reader on two related cp1250 samples. This is not exhaustive coverage of MS-OVBA variants.
- No VBA compilation, Excel runtime, live COM test or live-agent scenario execution is claimed for this release.

## 2.0.2 - 2026-10-25
- Redesigned the public GitHub README with a vector banner, Mermaid architecture and API-verification diagrams, a safer installation flow and explicit evidence boundaries.
- Reconciled public documentation, package layout, named release assets and checksum manifest; retained the v2.0.1 API verification features.
- Kept the separately distributed Microsoft documentation knowledge base unchanged.

## 2.0.1 - 2026-10-25
- Introduced read-only `api_preflight.py` for namespace-aware verification across VBA, Excel, MSForms, Office and VBIDE.
- Strengthened documentation selection, source ownership checks and regression coverage for unconfirmed MSForms members.
- Preserved offline Microsoft source files and the companion knowledge package.

## 2.0.0 - 2026-10-23
- Added an offline indexed knowledge package from the supplied Microsoft VBA-Docs snapshot: 38,187 Markdown documents with source integrity checks and full-source reads.
- Preserved existing modular engineering references, workflows and read-only VBA tools.
- Separated the core skill from the larger reusable knowledge archive.

## 1.0.0 - 2026-10-23
- Established Excel VBA Expert with domain-specific references, engineering workflows, read-only inspection tools, VBA examples and explicit evidence levels.
