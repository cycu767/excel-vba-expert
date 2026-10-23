# Tables, pivots and refresh integrity

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Tables
[DOC: MS-DATABODY] DataBodyRange represents table data without headers. [ENG] Define empty-table handling with a Nothing check, stable column names and explicit header/total-row treatment. Do not count a header or totals row as a business record. Validate schema before resizing or writing. Distinguish table formulas from constants, and decide whether filtered-out records belong to the operation.

Do not assume a table column position remains stable after a user inserts a column. Resolve ListColumns by name and fail closed on missing or duplicate business fields. Keep an idempotent row-key strategy for append/update workflows.

## Pivot cache decisions
[DOC: MS-PIVOT-CREATE] PivotCaches.Create supports particular source types; xlPivotTableVersionCurrent is not accepted. A qualified source string or named range can avoid some unexpected type-mismatch behavior. [ENG] Identify whether the source is worksheet data, external data or a model before selecting the API. Reuse a cache only when its sharing semantics are intended. Do not repeatedly create caches on every UI refresh without assessing workbook growth and cache ownership.

Validate destination, names, available fields and data types before manipulating a pivot. Refreshing cached data is not the same as modifying source records. Test filters and layout because a schema change can invalidate fields and calculated items.

## Refresh is not proof of fresh data
[DOC: MS-REFRESH] RefreshAll can return while background work remains; its current reference warns that a failing Power Query table refresh can leave old data without raising a runtime error. [DOC: MS-QUERY-REFRESH] QueryTable.Refresh with BackgroundQuery=False is synchronous for the documented supported query scenario.

[ENG] Record a refresh outcome per connection/table, not just a global call return. Check expected source version/timestamp, row count and domain reconciliation. A stale table may look valid. Do not set BackgroundQuery blindly on connection types that do not expose/support it. Handle unavailable credentials and prompts without logging secrets.

A timeout loop can bound polling but not an individual blocked COM call. Refreshing production connections can transmit queries, consume resources or change externally visible state; obtain permission. Keep test fixtures offline when live data is unnecessary.

## Tests
[ENG] Header-only table; renamed/reordered columns; mixed data types; zero rows; duplicates; expanded source; stale cache; failed query preserving old data; denied credentials; cancellation; asynchronous completion; repeat refresh without duplicate caches; unrelated workbook untouched.

## Source register

- [MS-DATABODY: ListObject.DataBodyRange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.databodyrange)
- [MS-PIVOT-CREATE: PivotCaches.Create](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcaches.create)
- [MS-REFRESH: Workbook.RefreshAll](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.refreshall)
- [MS-QUERY-REFRESH: QueryTable.Refresh](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.refresh)
