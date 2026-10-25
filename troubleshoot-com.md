# Diagnose external Excel automation safely

All steps below are engineering workflow rules, not claims that tests have already run.
## Identify the boundary
Record client (Python/VBS/other), pywin32 version if used, Windows/Office versions,
account/session, script entry, complete exception/HRESULT and last completed step.
Separate failure to create an application from failure to open a workbook, execute
VBA, refresh data or save. A COM exception and Err inside VBA are different evidence.

## Ownership and reproduction
Determine whether the client attached to an existing instance or created a dedicated
one. Only the owner may close/quit it. Use DispatchEx deliberately for isolation;
do not equate generic Dispatch with a guaranteed fresh process. Map workbooks,
child objects, scheduled work and modal dialogs. Do not kill all EXCEL.EXE processes.

Try the opt-in smoke test on a new blank workbook before reproducing with a sanitized
copy. Opening the user's file requires separate authorization, not just --run-local-
excel. Preserve Excel security policies; do not enable VBProject access or all macros.
Workbooks.Open with UpdateLinks=0 does not neutralize every data connection. Set and
restore AutomationSecurity appropriately for approved opening; ForceDisable has an
Excel 4.0 macro caveat and is not a general sandbox. Inspect active content first.

## Diagnose hangs and leaks
Expose the owned application when appropriate and look for save prompts, protected
view, password dialogs, hidden forms, VBA break mode, external refresh and add-ins.
Log timestamps around each blocking call. A Python timeout outside a synchronous
COM call does not necessarily interrupt that call; design a supervised worker only
when ownership, recovery and cancellation have been specified.

Release child ranges/charts/sheets and workbook references, close the owned workbook,
then attempt Quit even when Close failed. Preserve the original exception alongside
cleanup failures. Drop traceback-held references where necessary. Garbage collection
is not proof of process exit; separately observe the owned process when required.
Never terminate the user's unrelated Excel session to make a test green.

## Result
Give the failing boundary, ownership model, evidence-based cause or remaining
hypotheses, minimal change and test matrix. Distinguish mocked Python cleanup
coverage from real COM, VBA and UI tests. Retain sanitized diagnostics rather than
uploading workbooks or secrets to a remote service.
