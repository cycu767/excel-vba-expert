# Excel object model: ownership and API contracts

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Resolve ownership before a call
[DOC: MS-THISWORKBOOK, MS-ACTIVEWORKBOOK] The code container and active workbook are different concepts. ThisWorkbook is an internal VBA context, not an external Python selector. ActiveWorkbook can be Nothing. A menu action may deliberately target the active workbook; capture it once and validate it rather than banning it blindly.

[ENG] Establish a contract for every entry point: caller, workbook identity/full path, worksheet identity (Name versus CodeName), input ranges, outputs, save policy, Excel instance owner, and side effects. In an add-in or PERSONAL.XLSB, never replace a user's data workbook with ThisWorkbook merely to eliminate ActiveWorkbook.

Pass Workbook, Worksheet or Range explicitly into reusable logic. Resolve sheets by an agreed stable identifier, not a changing tab position. Use Worksheets when worksheet methods are required: Sheets may also include chart sheets. Preserve a worksheet CodeName used by existing event code.

```vb
Dim wbInput As Excel.Workbook
Dim wsInput As Excel.Worksheet
Set wbInput = Application.Workbooks.Open( _
    Filename:=inputPath, UpdateLinks:=0, ReadOnly:=True)
Set wsInput = wbInput.Worksheets("Input")
' Opening policy and cleanup belong to the caller; see security.md.
```
This is an ownership illustration, not a safe untrusted-file opener.

## Qualify the whole expression
[ENG] `ws.Range(Cells(1, 1), Cells(10, 3))` is unsafe in an ordinary module: its inner Cells can bind elsewhere. Use `ws.Range(ws.Cells(1, 1), ws.Cells(10, 3))` or a With block with a leading dot on EVERY member. Check qualifiers after copying a worksheet or showing a modeless form.

Do not mechanically remove Select/Activate when a documented operation requires activation. For routine reads/writes, work on the object directly. Treat Selection and ActiveCell as user-interface inputs requiring type, workbook and sheet validation.

## Match the library and object
[ENG] Resolve the declared type, runtime TypeName, library and version before choosing an API. ChartObject is an embedded container; Chart is its chart. Shape is not an MSForms control. Collection is not Scripting.Dictionary. Names can be workbook- or worksheet-scoped. A workbook connection is not automatically a QueryTable.

Lookup by fully qualified member (`Excel.Range.Find`), inspect signature/return value/applicability, then implement. Search misses in the partial local index do NOT prove a member does not exist. Escalate to a local repository index, installed Object Browser or permitted official documentation. Never fabricate a plausible member name.

## Acceptance probes
[ENG] Run with two workbooks open, switch active workbook immediately before invocation, and repeat from an add-in context where relevant. Assert source and destination full paths, sheet names, output addresses and absence of edits to unrelated workbooks. Include a chart-sheet-active and no-active-workbook case for UI entry points. API existence checks and successful compilation are separate from business-result tests.

## Source register

- [MS-OBJECT-MODEL: Excel object model](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)
- [MS-THISWORKBOOK: Application.ThisWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook)
- [MS-ACTIVEWORKBOOK: Application.ActiveWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeworkbook)
- [MS-OPEN: Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
- [MS-COLLECTION: Collection object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object)
- [MS-DICTIONARY: Dictionary object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)
