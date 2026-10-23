# Excel events: re-entry, boundaries and recovery

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Route the actual event
[DOC: MS-CHANGE] Worksheet.Change receives a Target that may contain multiple cells; recalculation is not a Change event. [ENG] Put sheet events in the correct sheet module, workbook events in ThisWorkbook, and application-event subscribers in a class with a retained WithEvents reference. A renamed public Sub in an ordinary module does not become an event handler.

Filter with `Application.Intersect(Target, Me.Range("B2:B1000"))`. Compare ranges on the same worksheet. Exit when there is no intersection. Define the policy for bulk paste, delete, formulas and errors rather than unconditionally discarding multi-cell changes.

## Guard only the necessary scope
[DOC: MS-EVENTS] EnableEvents controls Excel events. [ENG] Capture its entry value and restore that value on both success and failure. Disabling it is instance-wide, not limited to the current sheet. Avoid DoEvents inside this critical section. Do not assume it suppresses MSForms events; use a form-local loading/re-entry flag there.

Use `assets/vba/WorksheetChange.form-code.txt` as a sheet-module example and `MApplicationState.bas` for general state restoration. Merge into an existing event instead of adding a second procedure with the same signature. A silent error handler that leaves events disabled is a defect, not resilience.

## Design the event as an adapter
[ENG] Keep the event procedure short: qualify/filter input, reject unsupported state, invoke a domain procedure, restore state, then report the outcome. Business calculation should be callable without raising UI events. Document whether programmatic writes are expected to trigger downstream work; disabling all events may suppress essential validation.

For modeless forms, validate that the target workbook is still open before committing. For queued callbacks or OnTime work, record scheduling and cancellation ownership; do not leave a callback referencing an unloaded form. APIs outside the curated index need lookup before use.

## Diagnostic sequence
[ENG] First check the module placement and exact event signature, then `Application.EnableEvents`, Target.Address/CountLarge, and the sheet identity. Compile the project. Break at entry before assuming the handler is not firing. Check whether the apparent change is a recalculation. Capture the initiating user action and any nested writes; re-entry is a causal hypothesis until observed.

## Tests
[ENG] Single edit; multi-cell paste; deletion; irrelevant column; invalid value; protected target; deliberately raised domain error; original EnableEvents=False; repeated run. Verify that the state after each run equals the entry state. Force an error AFTER disabling events to exercise cleanup. Reset/End/process termination can bypass recovery; never promise guaranteed restoration after them.

## Source register

- [MS-CHANGE: Worksheet.Change](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.change)
- [MS-EVENTS: Application.EnableEvents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents)
- [MS-DIM: Dim statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
- [MS-DOEVENTS: DoEvents function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doevents-function)
- [MS-ON-ERROR: On Error statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement)
