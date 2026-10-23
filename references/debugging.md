# Evidence-led VBA debugging

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Start with a reproducible failure
[ENG] Record exact symptom, Err.Number/Description/Source, failing procedure/line, expected and actual output, workbook copy, input dimensions, Office build/bitness, references and entry point. Use the same input for before/after tests. A screenshot of a highlighted line is evidence of location, not proof of root cause. Do not silently change business rules to remove an error.

## Debugger toolkit
[DOC: MS-DEBUG-MENU, MS-TRACE, MS-BREAKPOINT, MS-IMMEDIATE]

| Operation | Purpose | Typical shortcut |
|---|---|---|
| Compile project | Catch syntax/type/reference errors | Debug menu |
| Toggle breakpoint | Halt before suspect action | F9 |
| Step Into | Inspect a callee | F8 |
| Step Over | Execute a call as a step | Shift+F8 |
| Step Out | Finish current procedure | Ctrl+Shift+F8 |
| Run to Cursor | Skip verified sections | Ctrl+F8 |
| Immediate | Query local state / Debug.Print | Ctrl+G |
| Quick Watch | Inspect a selected expression | Shift+F9 |

[ENG] Prefer read-only probes: `? TypeName(obj)`, workbook FullName, range Address, VarType and array bounds. Guard Nothing/error values before probing members. Immediate commands can change data; never run an unknown function merely to inspect it. Use Locals for current scope and Call Stack for the actual call chain. Limit watches to the relevant scope and remove stale watches.

[DOC: MS-ASSERT] Debug.Assert can suspend execution. [ENG] Keep assertions free of side effects, and use explicit error-raising assertions for unattended test harnesses so a failure does not hang in an unseen IDE.

## Isolate by failure class
[ENG] Compile/reference: inspect MISSING entries and actual library requirements; do not blindly uncheck references. Runtime 9: inspect collection owner/key/bounds. Runtime 13: inspect Variant subtype, array/scalar shape and Null/error cells. Runtime 91: trace where the object should have been assigned. Runtime 1004: reduce to the precise Excel call, qualified objects, protection and parameter state; it is not one universal error. These are diagnostic leads, not confirmed diagnoses.

For events, use events.md. For forms, examine instance lifetime, loading guards and modal windows. For COM, capture HRESULT/excepinfo and the owning process; inspect prompts in a controlled visible run before introducing retries. For hangs, distinguish an active calculation, blocked modal dialog, re-entry loop and a synchronous COM call that never returns.

## Stop safely
[DOC: MS-DEBUG-TOOLBAR, MS-STOP] Break and Reset are different. [ENG] Prefer a cooperative cancel or break to inspect state. Reset, End and termination may bypass cleanup. Record/restore known application settings under operator control; do not blindly force them to True/Automatic when the entry state is unknown. Never terminate every EXCEL.EXE to solve one test's orphan.

## Close the loop
[ENG] State a falsifiable hypothesis, collect confirming/disconfirming evidence, apply the smallest patch, rerun the reproducer and adjacent cases, compare outputs/state, and label untested steps. Do not replace the whole module merely because the first fix failed.

## Source register

- [MS-DEBUG-MENU: Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)
- [MS-DEBUG-TOOLBAR: Debug toolbar](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-toolbar)
- [MS-TRACE: Trace code execution](https://learn.microsoft.com/en-us/office/vba/language/how-to/trace-code-execution)
- [MS-BREAKPOINT: Set and clear a breakpoint](https://learn.microsoft.com/en-us/office/vba/language/how-to/set-and-clear-a-breakpoint)
- [MS-IMMEDIATE: Use the Immediate window](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/use-the-immediate-window)
- [MS-QUICK-WATCH: Use Quick Watch](https://learn.microsoft.com/en-us/office/vba/language/how-to/use-quick-watch)
- [MS-WATCH-ADD: Add a watch expression](https://learn.microsoft.com/en-us/office/vba/language/concepts/forms/add-a-watch-expression)
- [MS-WATCH-EDIT: Edit a watch expression](https://learn.microsoft.com/en-us/office/vba/language/how-to/edit-a-watch-expression)
- [MS-WATCH-DELETE: Delete a watch expression](https://learn.microsoft.com/en-us/office/vba/language/how-to/delete-a-watch-expression)
- [MS-ASSERT: Debug.Assert](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assert-method)
- [MS-STOP: Stop code execution](https://learn.microsoft.com/en-us/office/vba/language/how-to/stop-code-execution)
