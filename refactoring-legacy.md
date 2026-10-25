# Legacy VBA audit and low-risk modernization

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Build a baseline inventory
[ENG] Work on a copy under version control. Record file hash, format, VBA modules/classes/forms, named ranges, buttons/OnAction, sheet/workbook events, external links/connections, references/OCX, Windows API declarations, scheduled callbacks, UDFs and external automation entry points. Export sources through approved tooling; do not automatically enable Trust access to the VBA project object model.

Inventory .frm/.frx pairs and any custom UI resources. The read-only package inspector can identify ZIP parts, not recover every form layout or certify macro safety. Record inaccessible/encrypted/protected components as unknown rather than reporting them absent.

## Characterization tests first
[ENG] Capture the application's current behavior on sanitized fixtures: expected totals, output files, ordering, rounding, user messages, cancellation behavior and application settings. Distinguish a bug-preserving baseline from a deliberate approved correction. Large legacy business applications often encode important contracts in event order or worksheet names; cosmetic cleanup can break them.

Separate changes into independent slices: qualifier fix, error-state cleanup, extracted pure transformation, UI layout, library replacement, 64-bit API migration. Do not mix all of them into one rewritten module. Keep the old/new diff reviewable and a rollback path available.

## Dead code is a hypothesis
[ENG] No text references does not establish unused code. Search event signatures, Application.Run strings, Shape.OnAction, ribbon XML callbacks, named formulas, UDFs, macro dialogs, workbook links and external scripts. If an entry point cannot be exercised or traced, mark it candidate-unused. Do not delete it automatically.

## Dependency and architecture risks
[ENG] Map globals and shared instances to their writers/readers. Identify load-order dependence, implicit active workbook context, duplicated calculations, hidden error suppression, code that closes user workbooks, and references kept alive by modeless forms or event sinks. Refactor around stable interfaces, not around the assumption that every duplicate-looking formula is semantically identical.

[DOC: MS-64BIT] Pointer-width migration needs more than PtrSafe. [ENG] Audit declarations, storage, UDTs and optional ActiveX controls separately; successful compilation on one machine does not validate deployment to another.

## Report and release
[ENG] Deliver an evidence-ranked risk register, dependency map, preserved invariants, proposed small change batches and a test matrix. Label static suspicions, reproduced failures and unresolved items separately. Promote a change only after its targeted regressions and output reconciliation pass on the target Excel environment; otherwise keep it an unverified candidate patch.

## Source register

- [MS-64BIT: 64-bit VBA overview](https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview)
- [MS-THISWORKBOOK: Application.ThisWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook)
- [MS-ACTIVEWORKBOOK: Application.ActiveWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeworkbook)
- [MS-ON-ERROR: On Error statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)
- [MS-DEBUG-MENU: Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)
