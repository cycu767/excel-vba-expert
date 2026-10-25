# Performance without semantic drift

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Measure the boundary, not a guess
[ENG] Record workbook copy/hash, input size, Office build/bitness, calculation mode, cold/warm run, elapsed time and validated output totals. Profile data reads, transformations, writes, calculation, refresh, chart updates and save separately. A faster run with stale calculation or fewer records is not an optimization.

[DOC: MS-PERFORMANCE, MS-VALUE] Worksheet/VBA transfers and repeated UI work can dominate cost. [ENG] Read bounded rectangular blocks into arrays, transform in memory and write once when memory allows. Chunk genuinely large data sets rather than assuming an entire worksheet fits in memory. Avoid cell-by-cell COM calls from Python; keep explicit scalar/matrix normalization at that boundary.

## Controlled application settings
[DOC: MS-PERFORMANCE] Microsoft recommends restoring original settings after temporary changes. [ENG] Disable only what the operation requires. Manual calculation affects the Excel instance, not merely the output sheet. Decide what must be recalculated before validation/save and restore the previous mode. ScreenUpdating is a performance setting, not a security control. DisplayAlerts=False is not a standard optimization flag.

Do not call Calculate repeatedly inside a row loop unless each iteration logically depends on new calculated results. Avoid volatile functions and full-column array operations when a bounded range is sufficient, but do not rewrite formulas without regression tests. Profile lookup algorithms against duplicate-key and ordering rules; replacing a scan with a Dictionary can change case comparison or which duplicate wins.

## Re-entry and responsiveness
[DOC: MS-DOEVENTS] DoEvents processes queued events. [ENG] Throttle progress updates and use a re-entry guard with explicit cancellation checkpoints. Never yield while a half-written financial result is externally visible unless readers can distinguish an in-progress state. Do not wrap an infinite loop in DoEvents and call it responsive.

## Performance acceptance
[ENG] Compare median/variation over repeated runs on the same fixture. Confirm record counts, formulas, totals, key ordering, chart contents and application settings. Separate expected changes from floating-point or locale drift. Keep a baseline and an upper time/memory budget suited to the project. Do not promise a universal speedup percentage from using arrays or turning off calculation.

## Source register

- [MS-PERFORMANCE: Excel performance obstructions](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)
- [MS-VALUE: Range.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value)
- [MS-VALUE2: Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
- [MS-DOEVENTS: DoEvents function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doevents-function)
