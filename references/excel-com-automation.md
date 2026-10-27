# Excel COM automation and process ownership

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Separate client code from VBA
[DOC: MS-THISWORKBOOK] ThisWorkbook is not the way an external automation client identifies a workbook. [ENG] Python must retain the object returned by its own Workbooks.Open/Add call. A Python wrapper, a COM server and VBA inside that server are three separate layers with different exception and lifetime behavior.

Check Windows, installed desktop Excel and optional pywin32 before attempting COM. Do not install a package or alter registry/Trust Center settings automatically. A 64-bit Python process is not proof of 64-bit Office, and out-of-process COM should not be rejected solely because client/server bitness differs; verify actual dependencies and tests.

## Own what you create
[DOC: PYWIN32-DISPATCH] The primary pywin32 source distinguishes obtaining an active object from CoCreateInstanceEx-based creation. [ENG] Use DispatchEx for an isolated test instance, retain its identity and never call Quit on a borrowed interactive instance. Do not claim generic Dispatch always creates a new process. Use GetActiveObject only when the user explicitly chose an existing instance and agreed on close/save behavior.

Release child Range/Worksheet/Chart references before workbook/instance teardown. Close only workbooks the operation owns, with explicit SaveChanges. Attempt Quit in finally even if Close fails; record cleanup errors without hiding the primary exception. Avoid stale references retained in globals, collections, debugger variables and exception tracebacks. `gc.collect()` alone is not proof the Excel process exited.

## Security before opening
[DOC: MS-OPEN, MS-AUTOMATION-SECURITY] Programmatic opening needs explicit macro security, and ForceDisable does not cover Excel 4.0 macros. [ENG] Inspect untrusted files without Excel first. Use isolated, policy-approved environments for untrusted content; this skill provides no general safe untrusted-file opener. UpdateLinks=0, ReadOnly=True and disabling events are additional constraints, not a complete sandbox or a guarantee against connection refresh.

Keep macro execution and reading separate. Never lower macro security to run an unverified workbook. An approved Application.Run workflow must name a trusted workbook/macro explicitly, run a copy, record input/output, and respect organizational policy without enabling VBA-project access.

## Orphaned instances poison later runs
[ENG] A crashed or interrupted automation script can leave a hidden Office
instance alive. Later gates that attach by lookup rather than by creation then
connect to that corpse and fail en masse with an environment error such as
"macros have been disabled" or "cannot run the macro". That is not a product
regression. Enumerate and account for orphaned owned processes before a gate run,
and treat simultaneous identical failures across unrelated gates as an
environment signal that must be diagnosed first.

[ENG] Distinguish attaching from creating in the gate itself, and record which
instance identity each gate used. Before trusting a failure, check how many
workbooks the attached instance holds and whether it is the one this run created.
Never resolve this by killing every Office process on the machine.

## COM calls and waits
[ENG] Capture pywintypes.com_error HRESULT, excepinfo, operation name and object identity, with sensitive values redacted. Inspect modal dialogs in a supervised visible run. Retry only classified transient rejection with a finite budget; do not retry a non-idempotent operation blindly. A loop deadline does not interrupt a single blocked COM call. Hard timeouts require a separately owned worker and a policy for that worker's Excel instance, not taskkill of all Excel processes.

If using threads, initialize/uninitialize COM on the calling thread and respect apartment rules. Do not pass raw Excel COM proxies between arbitrary threads. This skill does not implement a production process supervisor or unattended server automation platform.

## Included smoke test
[ENG] `scripts/excel_smoke_test.py --run-local-excel` deliberately creates a NEW blank workbook, writes test data/formula, creates a chart, saves to a NEW temporary directory, then closes and quits its owned instance. It accepts no user workbook path and runs no user macro. Default invocation performs dependency checks only. Its Python control flow is mock-tested; live COM execution remains a target-Windows acceptance test.

## Source register

- [MS-THISWORKBOOK: Application.ThisWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thisworkbook)
- [PYWIN32-DISPATCH: pywin32 client source](https://github.com/mhammond/pywin32/blob/main/com/win32com/client/__init__.py)
- [MS-OPEN: Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
- [MS-AUTOMATION-SECURITY: Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
- [MS-QUIT: Application.Quit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.quit)
- [MS-ALERTS: Application.DisplayAlerts](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts)
