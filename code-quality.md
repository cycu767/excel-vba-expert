# Code review standards and actionable findings

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Review contracts before style
[ENG] Identify entry points, declared side effects, workbook ownership, dependencies and expected invariants. Review in order: data-loss risk, security, result correctness, state/lifetime recovery, testability, performance, then formatting. Do not add a class hierarchy or service abstraction just to imitate another language.

Require readable procedure names, clear inputs/outputs, explicit type/scope and small cohesive responsibilities. Avoid magic sheet/column indexes when the data schema has stable names. Keep explanatory comments about constraints and rationale, not narration of every assignment. Preserve public signatures unless compatibility changes are explicitly approved.

## Candidate rules
[ENG] Flag unqualified Range/Cells/Sheets in standard modules; active workbook/selection reliance; broad On Error Resume Next; End; unchecked Find results; state changes with no visible cleanup; implicit Variant declarations; unguarded event writes; unbounded waits; unowned Quit/Close; destructive file operations and external execution. Flag does not mean proven defect: event/class module context and required activation can justify an occurrence.

Use `scripts/inspect_vba.py` only as a read-only heuristic assistant. It strips ordinary strings/comments but is not a VBA parser, resolver or compiler. It does not prove dead code, conditional-compilation correctness, complete call graphs or absence of vulnerabilities. Review each finding in context and state confidence.

## Record evidence
[ENG] For each issue produce: ID; severity; confirmed/hypothesis; file/procedure/line; minimal relevant excerpt; trigger; impact; recommended smallest change; regression test; residual uncertainty. Do not call code unreachable without examining event hooks, OnAction, Application.Run, ribbon callbacks, named formula/UDF references and external callers.

## Export/import discipline
[ENG] Preserve .cls attributes and document modules; importing an exported sheet class as an ordinary class does not rebind it to the worksheet. Treat .frm/.frx as a pair. Check actual text encoding before normalizing files; localized strings can be corrupted by an ANSI/UTF-8 round trip. The bundled VBA templates use ASCII to avoid that specific transport issue, not to restrict project-language support.

## Quality gate
[ENG] Compile in VBE, run targeted tests, verify business outputs and inspect the diff for unrelated changes. If those capabilities are unavailable, deliver a patch plus a concrete test plan and label the evidence as static only. A lint report with no findings is not a release certificate.

## Source register

- [MS-EXPLICIT: Option Explicit statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement)
- [MS-DIM: Dim statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
- [MS-ON-ERROR: On Error statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)
- [MS-FIND: Range.Find](https://learn.microsoft.com/en-us/office/vba/api/excel.range.find)
- [MS-DEBUG-MENU: Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)
- [MS-THISWORKBOOK: Application.ThisWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook)
