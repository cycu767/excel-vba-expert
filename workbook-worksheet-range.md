# Workbook, worksheet and range operations

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Range shape is a precondition
[DOC: MS-VALUE] Reading multiple cells gives a two-dimensional Variant array, whereas a scalar cell must be treated separately. Array writes to multi-area ranges are not a reliable general operation. [ENG] Validate Areas.Count, rows/columns, merged cells, sheet protection and destination boundaries before a bulk write. Normalize a scalar to a 1x1 array only when the procedure contract requires an array.

[DOC: MS-VALUE2] Value2 does not preserve Date/Currency subtypes. [ENG] Choose Value2 for raw block transfer, but define date and monetary conversions explicitly. Preserve error values intentionally or reject them; do not convert every Variant with CStr. A value transfer overwrites destination formulas. State that effect before applying it.

## Reliable lookup
[DOC: MS-FIND] Find can return Nothing, and its options can persist. [ENG] Define What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte and SearchFormat deliberately. After must be a single cell inside the search range. For a repeated search, save the first address and stop on wraparound. Define whether `*`, `?` and `~` are wildcards or literal input.

```vb
Dim hit As Excel.Range
Set hit = searchRange.Find(What:=key, After:=searchRange.Cells(1, 1), _
    LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False, _
    SearchFormat:=False)
If hit Is Nothing Then
    Err.Raise vbObjectError + 2101, "FindKey", "Key not found"
End If
```
[ENG] This pattern assumes `key` is already validated and a literal/wildcard policy was chosen. It changes Find settings visible in the UI; do not describe it as zero-side-effect.

## Extents, names, sorting and filtering
[ENG] Select last-row logic using the dataset contract. End(xlUp) on a required key column does not define the extent of unrelated columns. UsedRange can include formatting and is not a business-record count. Account for formula cells displaying empty strings. Handle zero rows without creating an inverted range.

Resolve named ranges in their intended workbook/sheet scope; a name may refer to a formula rather than cells. Use table column names where stable, not magic indexes. Sort all columns that constitute a record and explicitly declare header presence. With filters, distinguish all records from visible records; do not assume a filtered range is contiguous. Verify zero visible records before copying. Reapply validation deliberately when replacing a range.

## Workbook lifecycle
[DOC: MS-OPEN, MS-SAVECOPY] Capture the Workbook returned by Open. SaveCopyAs writes a copy without altering the open workbook. [ENG] Do not close a workbook borrowed from the user. Preserve file format and macro content; saving a VBA project as xlsx is not a harmless rename. Use new output paths, detect collisions, and reopen a saved test output before promoting it. A backup is not rollback of in-memory writes.

## Required tests
[ENG] Cover no data, one cell, one row, multi-area selection, error cells, duplicate keys, merged/protected destinations, hidden rows, header-only tables, names with different scopes, and two workbooks with identically named sheets. Compare formulas as well as values when formulas are in scope.

## Source register

- [MS-VALUE: Range.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value)
- [MS-VALUE2: Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
- [MS-FIND: Range.Find](https://learn.microsoft.com/en-us/office/vba/api/excel.range.find)
- [MS-OPEN: Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
- [MS-SAVECOPY: Workbook.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas)
- [MS-DATABODY: ListObject.DataBodyRange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.databodyrange)
