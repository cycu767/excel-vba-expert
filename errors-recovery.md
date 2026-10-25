# Error handling and application-state restoration

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Distinguish error states
[DOC: MS-ON-ERROR, MS-RESUME] An enabled handler and an active handler are not the same. A new error while the current handler is active propagates to a caller. Resume is not equivalent to an arbitrary GoTo out of the handler. [ENG] Capture original Err.Number, Source and Description before cleanup can overwrite them. If Erl is used, remember it only provides useful source locations where line numbering exists; do not report zero as a real failing line.

## The recovery contract
[ENG] Separate preflight validation, capture of state, mutation, cleanup, and propagation. Capture all needed properties BEFORE changing any; mark the snapshot valid only after capture succeeds. Restore exact old values, including EnableEvents=False or manual calculation already present on entry. Restore calculation while Excel events remain disabled, then restore other settings and events last. Document that recalculation can have consequences and avoid changing calculation unless necessary.

`assets/vba/MApplicationState.bas` provides a complete module for copying a rectangular block of values with explicit source/target objects and error-preserving cleanup. It does not claim transactionality or successful VBE compilation in this build.

## Local suppression only
[ENG] Permit On Error Resume Next only around a narrowly defined probe or cleanup operation. Clear Err immediately before, capture it immediately after, turn suppression off, and restore the intended handler if needed. Do not use broad suppression to skip missing data, silently drop failed rows, or claim a failed save succeeded.

In a cleanup routine, attempt every independent restoration even if one fails; retain the first cleanup failure and include it in the outcome. If there was an original business failure, propagate that first with cleanup diagnostics. If only cleanup failed, report that failure. Never emit an unconditional success message at the end of a catch-all handler.

## Retry and idempotency
[ENG] Retry only a classified transient condition with a bounded count and a deadline. A retry of append/save/send can duplicate side effects. Define a transaction key or detect whether the operation already committed. Resume on a permanent invalid argument simply repeats the defect. Keep recovery from a file-write failure separate from recovery of Excel application settings.

## User interruption
[ENG] Do not use End for normal control flow. A cooperative cancellation should pass through cleanup. VBE Reset, host crashes, power loss and external termination cannot be covered by an in-process error handler. Keep a durable operation marker where recovery requirements justify it. After an uncertain interruption, fail closed rather than promoting potentially partial output.

## Acceptance checks
[ENG] Inject failure before capture, immediately after disabling events, during data write, during save, and during each independent cleanup step. Assert original error preservation and final state. Include nested callers where the parent intentionally disabled events. Mock tests can validate control flow in Python, but cannot certify VBA event or COM behavior.

## Source register

- [MS-ON-ERROR: On Error statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)
- [MS-RESUME: Resume statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/resume-statement)
- [MS-EVENTS: Application.EnableEvents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents)
- [MS-ALERTS: Application.DisplayAlerts](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts)
- [MS-ACCESS-ERRORS: Elements of run-time error handling (Access)](https://learn.microsoft.com/en-us/office/vba/access/concepts/error-codes/elements-of-run-time-error-handling)
- [MS-PERFORMANCE: Excel performance obstructions](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)
