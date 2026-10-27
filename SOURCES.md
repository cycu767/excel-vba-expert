# Sources and measured offline coverage - version 2.0.0

## Source snapshot and identity
Imported the actual user-supplied `VBA-Docs-main.zip`, without a network download.
Repository attributed by the supplied archive: MicrosoftDocs/VBA-Docs.
Archive SHA-256: `70ca62b528225895836053deecf65458a42ef9f9cc808ab007060a4e50a1afa6`.
Archive comment: `6765796079a1552460769cb633068d387594dc99`. This has the form of a Git
commit SHA; its origin was not independently authenticated over the network.
The source files are preserved byte-for-byte, with no silently corrected examples.

## Measured scope
| Measure | Actual value |
|---|---:|
| Source files in supplied archive | 38,911 |
| Source bytes (logical size) | 61,172,477 |
| Markdown files physically imported | 38,187 |
| Full-text indexed Markdown files | 38,187 |
| Local Markdown source files readable | 38,187 |
| Excel-prefixed API Markdown files | 5,997 |
| Excel concepts | 108 |
| Office API Markdown files | 1,536 |
| Language tree Markdown files | 1,492 |
| Forms-concepts subset of Language | 82 |
| Preserved expert domain modules | 19 |
| Preserved workflows | 7 |
| Preserved original Python tool/helper files | 8 |
| Curated summary records | 70 |
| Curated IDs mapped to complete local Markdown | 62 |

Every supplied Markdown document is indexed, including navigation, F1, legacy Graph
and other Office products. **This is 100% mechanical coverage of the supplied
snapshot, not proof of all current Microsoft documentation or all possible APIs.**
The 5,997 Excel API paths include navigation/legacy pages; do not call them 5,997
objects. Full source availability is separate from authored guidance and runtime
verification. The domain modules still represent selected expert guidance.

## Local source architecture
The separate knowledge package has `repository/` (original full text and assets),
`index.sqlite` (metadata, exact aliases, full-text index), `documents.jsonl` (local
source locations and metadata), `files.jsonl` (source hashes), `manifest.json`
(snapshot/counts/index hashes) and `curated-crosswalk.json` (two-layer integration).
All document locations are relative to the knowledge root. CLI commands resolve
the root relative to the installed skill, independent of the project directory.

## Disjoint path-category inventory
| Category | Indexed Markdown |
|---|---:|
| excel-api | 5,997 |
| excel-concepts | 108 |
| excel-overview | 4 |
| f1-index | 2,118 |
| forms-concepts | 82 |
| office-api | 1,536 |
| office-concepts | 12 |
| office-mac | 6 |
| other-access | 1,065 |
| other-api-access | 3,703 |
| other-api-outlook | 4,521 |
| other-api-powerpoint | 2,812 |
| other-api-powerpoint(enumerations) | 1 |
| other-api-project | 3,263 |
| other-api-project(enumerations) | 1 |
| other-api-publisher | 1,702 |
| other-api-publisher(enumerations) | 1 |
| other-api-toc | 1 |
| other-api-visio | 2,809 |
| other-api-visio(enumerations) | 1 |
| other-api-word | 6,226 |
| other-api-word(enumerations) | 1 |
| other-api-word(textcolumns) | 1 |
| other-articles | 1 |
| other-outlook | 458 |
| other-powerpoint | 23 |
| other-project | 9 |
| other-publisher | 4 |
| other-visio | 24 |
| other-word | 74 |
| overview | 208 |
| repository-meta | 3 |
| support | 2 |
| vba-concepts | 52 |
| vba-how-to | 28 |
| vba-reference | 1,324 |
| vbe | 6 |

The Forms reference is interleaved with VBA language/UI reference in the original
repository. The 82 Forms-concept documents are **not** the total Forms documentation.
The vbe path category contains five add-in-model documents plus its UI-help
reference page, all inside Language; the Language total is counted independently
from the original archive paths.

## Known source/documentation gaps
All files in the supplied archive are present. Links to pages outside this repository
remain external; the importer does not fetch them. Of the eight curated IDs without
a Markdown crosswalk, MS-REPO, MS-DOCFX, MS-LICENSE and MS-LICENSE-CODE concern the
repository or non-Markdown files, which are locally preserved. VSC-SKILLS,
MS-GETACTIVEWINDOW, MS-SLEEP and PYWIN32-DISPATCH refer to external documentation;
their existing summaries remain, not full offline copies of those external sites.

The ChartObjects.Add source lists Left, Top, Width, Height in syntax but only Left
and Width in its parameter table, and no Remarks section. UserForm.Show is the
source page titled Show method with a parts table in Syntax. These are recorded
as source facts rather than repaired silently. See [API verification](OFFLINE-VERIFICATION.md).

## Licensing and semantic review
Microsoft and contributors' original LICENSE, LICENSE-CODE and ThirdPartyNotices
are included unchanged in the knowledge repository and copied into skill licenses/.
The raw source has not been semantically reviewed page-by-page. The reviewed core
register below is retained from v1; dates describe that review, not a claim that
all repository pages were reviewed or updated on the build date.

## Reviewed source register
The following summaries and domain guidance are original paraphrases/engineering
synthesis. Source authors retain their rights. No Microsoft endorsement is implied.

### MS-EXCEL - Excel VBA reference

[https://learn.microsoft.com/en-us/office/vba/api/overview/excel](https://learn.microsoft.com/en-us/office/vba/api/overview/excel)

Entry point for Excel concepts, the Excel object model and the separate Graph Visual Basic reference.

### MS-OBJECT-MODEL - Excel object model

[https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)

Navigation entry for objects, methods, properties, events and enumerations. This landing page is not an inventory of every member.

### MS-GRAPH - Excel Graph Visual Basic reference

[https://learn.microsoft.com/en-us/office/vba/api/overview/excel/graph-visual-basic-reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/graph-visual-basic-reference)

Separate Graph object-model documentation, not Microsoft Graph REST and not an algorithmic dependency graph.

### MS-LANGUAGE - VBA language reference

[https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference](https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference)

Connects language concepts, Microsoft Forms, how-to guidance, language members, definitions, editor add-in model and VBE UI help.

### MS-LIBRARY - Office library reference

[https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference](https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference)

Entry to Office-wide library concepts and object-model reference.

### MS-REPO - VBA-Docs repository

[https://github.com/MicrosoftDocs/VBA-Docs](https://github.com/MicrosoftDocs/VBA-Docs)

Public Microsoft documentation repository with Language, Library-Reference, api and Excel concept branches. Includes documentation and code licenses.

### MS-DOCFX - VBA-Docs publishing configuration

[https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/docfx.json](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/docfx.json)

DocFX includes Markdown and YAML, excludes include files and license files from article publishing, and assigns metadata by paths. Path inference must still distinguish publishing redirects and case.

### MS-LICENSE - VBA-Docs documentation license

[https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/LICENSE](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/LICENSE)

CC BY 4.0 license: preserve attribution and notices, identify changes, and provide the license link. No endorsement is granted.

### MS-LICENSE-CODE - VBA-Docs code license

[https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/LICENSE-CODE](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/LICENSE-CODE)

Microsoft code license is MIT; retain copyright and permission notice when redistributing covered code.

### VSC-SKILLS - VS Code Agent Skills

[https://code.visualstudio.com/docs/agent-customization/agent-skills](https://code.visualstudio.com/docs/agent-customization/agent-skills)

Personal skills may be placed in ~/.copilot/skills/. Instructions are loaded on demand. Discovery does not guarantee invocation; support depends on the agent harness.

### MS-OPEN - Workbooks.Open

[https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)

Returns the opened Workbook. Explicitly choose link updates and read-only mode. Programmatic opening requires an explicit macro-security policy.

### MS-AUTOMATION-SECURITY - Application.AutomationSecurity

[https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)

ForceDisable controls macros in programmatically opened files but does not disable Excel 4.0 macros. DisplayAlerts does not suppress security warnings. Capture and restore the prior security state.

### MS-CHANGE - Worksheet.Change

[https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.change](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.change)

Target can contain multiple cells. Recalculation changes do not trigger this event; use Calculate for recalculation-specific work.

### MS-FIND - Range.Find

[https://learn.microsoft.com/en-us/office/vba/api/excel.range.find](https://learn.microsoft.com/en-us/office/vba/api/excel.range.find)

Returns Nothing when no match exists. Several search options persist between calls and the Find dialog; specify them explicitly. Handle wraparound in repeated searches.

### MS-VALUE2 - Range.Value2

[https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)

Value2 does not use the Currency or Date subtypes; those values can be exposed as Double. It is not a formatting or locale parser.

### MS-VALUE - Range.Value

[https://learn.microsoft.com/en-us/office/vba/api/excel.range.value](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value)

Multi-cell first-area reads return a two-dimensional Variant array. Multi-area array assignment is not properly supported. Size the destination to the array.

### MS-THISWORKBOOK - Application.ThisWorkbook

[https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook)

Identifies the workbook containing the running macro. It is not an external COM-client workbook selector. Add-in code and the active user workbook may differ.

### MS-ACTIVEWORKBOOK - Application.ActiveWorkbook

[https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeworkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeworkbook)

Identifies the workbook in the active window and can be Nothing. Protected View has separate access semantics.

### MS-FORMULA2 - Range.Formula2

[https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2)

Uses array-evaluation semantics and may spill. Formula remains for backward compatibility. Confirm the target Excel supports dynamic arrays.

### MS-DATABODY - ListObject.DataBodyRange

[https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.databodyrange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.databodyrange)

Identifies the table data range excluding its header. The skill additionally requires a defensive Nothing check for empty-table workflows.

### MS-ALERTS - Application.DisplayAlerts

[https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts)

Suppresses selected prompts, not security. SaveAs may overwrite without confirmation. Cross-process callers must restore state explicitly.

### MS-ON-ERROR - On Error statement

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)

Distinguish enabled from actively handling error handlers. An error inside an active handler propagates to its caller. Resume Next must have a narrow justified scope.

### MS-RESUME - Resume statement

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/resume-statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/resume-statement)

Resume retries the failing statement, Resume Next continues after it, and Resume label resumes at a local label. It requires an active error context.

### MS-ASSERT - Debug.Assert

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assert-method](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assert-method)

Suspends execution when its expression is False. The whole boolean expression is evaluated; do not use it instead of production validation.

### MS-DEBUG-MENU - Debug menu

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)

Describes Compile, watch commands, breakpoints and stepping. Compile is a VBE operation, not equivalent to a text linter.

### MS-DEBUG-TOOLBAR - Debug toolbar

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-toolbar](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-toolbar)

Provides Run, Break, Reset, stepping and diagnostic windows. Reset clears the stack and module state.

### MS-BREAKPOINT - Set and clear a breakpoint

[https://learn.microsoft.com/en-us/office/vba/language/how-to/set-and-clear-a-breakpoint](https://learn.microsoft.com/en-us/office/vba/language/how-to/set-and-clear-a-breakpoint)

Use F9 to stop at a statement and Ctrl+Shift+F9 to clear all breakpoints. Breakpoints are not persisted with saved code.

### MS-STOP - Stop code execution

[https://learn.microsoft.com/en-us/office/vba/language/how-to/stop-code-execution](https://learn.microsoft.com/en-us/office/vba/language/how-to/stop-code-execution)

Distinguishes break mode from Reset and describes runtime errors, Stop, End and watch-triggered breaks.

### MS-TRACE - Trace code execution

[https://learn.microsoft.com/en-us/office/vba/language/how-to/trace-code-execution](https://learn.microsoft.com/en-us/office/vba/language/how-to/trace-code-execution)

Step Into enters callees; Step Over executes the call; Step Out completes the current procedure; Run to Cursor continues to a chosen statement.

### MS-IMMEDIATE - Use the Immediate window

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/use-the-immediate-window](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/use-the-immediate-window)

Ctrl+G opens Immediate. Commands execute in a module context and can mutate state; inspect values before using assignments or procedure calls.

### MS-ACCESS-ERRORS - Elements of run-time error handling (Access)

[https://learn.microsoft.com/en-us/office/vba/access/concepts/error-codes/elements-of-run-time-error-handling](https://learn.microsoft.com/en-us/office/vba/access/concepts/error-codes/elements-of-run-time-error-handling)

Supplementary explanation of error-handling flow. Transfer only language-level ideas; Access-specific object behavior is not an Excel contract.

### MS-WATCH-ADD - Add a watch expression

[https://learn.microsoft.com/en-us/office/vba/language/concepts/forms/add-a-watch-expression](https://learn.microsoft.com/en-us/office/vba/language/concepts/forms/add-a-watch-expression)

Define a watch and choose the narrowest relevant procedure/module context. Broad evaluation can slow execution.

### MS-WATCH-EDIT - Edit a watch expression

[https://learn.microsoft.com/en-us/office/vba/language/how-to/edit-a-watch-expression](https://learn.microsoft.com/en-us/office/vba/language/how-to/edit-a-watch-expression)

Edit the expression, scope and watch type in the Watch window.

### MS-WATCH-DELETE - Delete a watch expression

[https://learn.microsoft.com/en-us/office/vba/language/how-to/delete-a-watch-expression](https://learn.microsoft.com/en-us/office/vba/language/how-to/delete-a-watch-expression)

Remove an obsolete watch from the Watch window.

### MS-QUICK-WATCH - Use Quick Watch

[https://learn.microsoft.com/en-us/office/vba/language/how-to/use-quick-watch](https://learn.microsoft.com/en-us/office/vba/language/how-to/use-quick-watch)

Inspect a selected expression in break mode using Shift+F9 and optionally promote it to a watch.

### MS-USERFORM - UserForm object

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/userform-object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/userform-object)

UserForms represent loaded dialog instances. Multiple explicit instances may be created; controls are exposed through Controls.

### MS-INITIALIZE - Initialize event

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event)

Occurs after loading an object and before it is shown; use for initial preparation rather than repeated business actions.

### MS-ACTIVATE - Activate and Deactivate events

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/activate-deactivate-events](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/activate-deactivate-events)

Activation concerns visible active windows within the application. It is not a one-time constructor, and unloading does not produce Deactivate.

### MS-QUERYCLOSE - QueryClose event

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/queryclose-event](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/queryclose-event)

CloseMode identifies control-menu close, code unload and other close causes. Cancel can prevent closing; avoid blocking application shutdown accidentally.

### MS-HIDE - Hide method

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/hide-method](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/hide-method)

Hides without unloading; hidden controls remain accessible programmatically. Hiding an unloaded form can load it.

### MS-UNLOAD - Unload statement

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unload-statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unload-statement)

Unloads the form. Read a modal dialog result before unloading, rather than re-accessing a default instance afterward.

### MS-SHOW - Show method

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/show-method](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/show-method)

Supports modal and modeless display. Show loads an unloaded form. Code after modeless Show continues immediately.

### MS-CONTROLS-ADD - Add method (Microsoft Forms)

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms)

Adds tabs/pages or creates controls with a registered ProgID. Event wiring and ownership are separate engineering responsibilities.

### MS-CONTROLS - Controls collection

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/controls-collection-microsoft-forms](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/controls-collection-microsoft-forms)

Access controls by their names or zero-based collection indexes; enumerate each relevant container instead of equating index with tab order.

### MS-TABINDEX - TabIndex property

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tabindex-property](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tabindex-property)

Defines tab order, starting at zero. Verify focus order in the real form and its containers.

### MS-MULTIPAGE - MultiPage control

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/multipage-control](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/multipage-control)

Hosts a Pages collection; Value identifies the active page. Use distinct pages for groups of controls.

### MS-TABSTRIP - TabStrip control

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tabstrip-control](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tabstrip-control)

Hosts Tabs and exposes SelectedItem. Switching a tab can select different data for shared controls; it is not a MultiPage replacement by name alone.

### MS-64BIT - 64-bit VBA overview

[https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview](https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview)

PtrSafe alone does not fix truncated handles. Use LongPtr for pointers/handles and review UDTs and assignments; Long remains 32-bit.

### MS-PERFORMANCE - Excel performance obstructions

[https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)

Reduce unnecessary UI updates and worksheet/VBA transfers. Capture original application settings and restore them after work. Measure actual workloads.

### MS-CHART-ADD - ChartObjects.Add

[https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.add](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.add)

Creates an embedded ChartObject. Position and dimensions use points, not pixels.

### MS-CHART-SOURCE - Chart.SetSourceData

[https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setsourcedata](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.setsourcedata)

Sets chart source from an explicit Range with optional row/column plotting orientation.

### MS-NEWSERIES - SeriesCollection.NewSeries

[https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.newseries](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.newseries)

Creates a Series for supported charts; it is not available for PivotChart reports.

### MS-CHART-EXPORT - Chart.Export

[https://learn.microsoft.com/en-us/office/vba/api/excel.chart.export](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.export)

Exports using an installed graphic filter and returns Boolean. Check both result and the expected file instead of assuming export succeeded.

### MS-REFRESH - Workbook.RefreshAll

[https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.refreshall](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.refreshall)

Refresh may run in the background. The current page also warns that a failing Power Query table refresh can leave old data without RefreshAll raising a runtime error.

### MS-QUERY-REFRESH - QueryTable.Refresh

[https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.refresh](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.refresh)

Refreshes external data; BackgroundQuery=False makes supported SQL-query refresh synchronous. Missing connection details can produce dialogs or failures.

### MS-PIVOT-CREATE - PivotCaches.Create

[https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcaches.create](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcaches.create)

Creates a PivotCache from a supported source type. Explicit source strings/names can avoid unexpected type-mismatch issues. xlPivotTableVersionCurrent is not accepted.

### MS-QUIT - Application.Quit

[https://learn.microsoft.com/en-us/office/vba/api/excel.application.quit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.quit)

Quits Excel. Unsaved workbooks may prompt; suppressing alerts may discard them. Quit only the instance owned by the current operation.

### MS-EVENTS - Application.EnableEvents

[https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents)

Controls Excel application events. Preserve the previous value; do not assume it was True on entry.

### MS-SAVECOPY - Workbook.SaveCopyAs

[https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas)

Writes a copy without changing the open workbook in memory. It does not itself create a transactional rollback strategy.

### MS-EXPLICIT - Option Explicit statement

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement)

Requires variable declarations and surfaces undeclared names at compile time. It does not replace explicit typing or business validation.

### MS-IIF - IIf function

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/iif-function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/iif-function)

Evaluates both result expressions. Use an If block when a branch could fail or have side effects.

### MS-DIM - Dim statement

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)

Declares variables and arrays. WithEvents belongs in class modules, cannot form an array and cannot be combined with New.

### MS-SUB - Sub statement

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)

Declares procedures and parameters. VBA defaults to ByRef, unlike VB.NET. Specify parameter intent explicitly.

### MS-DOEVENTS - DoEvents function

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doevents-function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doevents-function)

Yields to the operating system to process queued events. Guard against re-entry; it does not turn VBA into a background worker.

### MS-TYPES - Data type summary

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary)

Summarizes VBA storage and ranges: Integer is 16-bit, Long 32-bit, Currency is scaled integer, and Double uses floating point.

### MS-COLLECTION - Collection object

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object)

Ordered group with Add, Item, Remove and Count. Do not assume it offers Dictionary.Exists.

### MS-DICTIONARY - Dictionary object

[https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)

Key/item container with Exists, Keys, Items and CompareMode. Availability of the scripting component is a deployment dependency.

### MS-GETACTIVEWINDOW - GetActiveWindow (Win32)

[https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getactivewindow](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getactivewindow)

Returns an HWND for the calling thread or NULL, not necessarily the foreground window of another process.

### MS-SLEEP - Sleep (Win32)

[https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-sleep](https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-sleep)

Suspends the current thread; its milliseconds parameter is DWORD, not a pointer. Sleeping on the UI thread is not a responsiveness solution.

### PYWIN32-DISPATCH - pywin32 client source

[https://github.com/mhammond/pywin32/blob/main/com/win32com/client/__init__.py](https://github.com/mhammond/pywin32/blob/main/com/win32com/client/__init__.py)

Primary implementation distinguishes GetActiveObject, Dispatch and DispatchEx. DispatchEx creates via CoCreateInstanceEx. Verify instance ownership rather than relying on generic Dispatch.
