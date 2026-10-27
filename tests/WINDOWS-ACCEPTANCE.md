# Windows / Excel / agent acceptance: NOT RUN in this build

## Preconditions
Use an approved Windows desktop with real Excel, a disposable directory and no
production workbooks in the test session. Record Office version/build/bitness,
Windows version, policy, references, locale and display scaling. Do not bypass
enterprise macro controls or change VBProject trust. An add-in-free test profile is
preferable: starting a separate application is NOT a sandbox against startup add-ins.
Do not open the user's workbook as part of an implicit smoke test.

## A. Skill discovery and actual agent use
Install the folder per README. In Copilot, use Chat: Open Customizations -> Skills;
confirm excel-vba-expert is listed. Start a fresh agent conversation and explicitly
ask it to use this skill. Check tool traces show SKILL.md then the correct workflow
and references, not all modules. Run all twenty scenarios in scenarios.json with the
actual model/client. Score each must_demonstrate and must_not_do criterion, save
model/client versions, prompts, responses and tool traces. A matching metadata route
is NOT an observed agent-routing pass. Qwen compatibility depends on its harness.

## B. Compile and pure VBA tests
Create a blank disposable .xlsm using Excel. Import MValidation.bas and
MValidationTests.bas, run Debug > Compile VBAProject, then RunValidationTests.
Expected: 10 assertions complete with PASS in Immediate, no error. Add other .bas
modules individually and compile. For MSForms create a real blank UserForm and import
CButtonHandler.cls; paste DynamicButtons.form-code.txt into its code module and
compile. No fabricated .frx is supplied. Repeat Windows API compilation on required
Office bitnesses; checking only the active #If branch is not proof for the other one.

## C. Workbook and failure tests
Create two disposable workbooks with distinct markers and explicit range variables.
Call CopyValuesSafely with events initially True, then False; calculation automatic,
then manual. Repeat with another workbook active. Verify marker values, no unrelated
writes and exact settings restoration. Reject mismatched dimensions before any write.
Inject a controlled Err.Raise after events are disabled in a COPY of the template;
verify original number/description and state. Test protected target and merged-cell
preconditions as separate cases. Cleanup cannot undo partially committed cell writes.

Install the worksheet example only in a test sheet, merge with no existing event.
Test single edit, multi-cell paste, a formula, an error cell and an out-of-scope edit.
Verify uppercase applies only to literal strings in the example range and no recursion.
For the chart module test repeated call, another unrelated chart, empty/bad dimensions
and intended data. Verify no unrelated chart is deleted.

## D. Forms
Show the dynamic-button form and click each button after initialization has returned.
Verify retained handler output. Close, reopen, switch workbooks, test keyboard and
repeat activation. A real production dialog additionally requires validation, Cancel,
X, Escape, Enter, modeless ownership and pending-edit policies. Check common target
DPI values, long messages, tab order, controls inside Frames/Pages and missing resources.

## E. COM smoke test
First run `python scripts/excel_smoke_test.py` for dependency status only. If Python,
pywin32 and Excel are installed and the user authorizes launching a new instance:
`python scripts/excel_smoke_test.py --run-local-excel --out <new-report.json>`.
Expected: new blank workbook, D1=8, one chart series, new .xlsx in a unique temp folder,
Close and Quit attempted. This does not run VBA, open existing files, inspect process
exit or test every UI path. Keep the output as evidence. Check manually that the owned
instance exits and unrelated user instances remain. The script cannot forcibly cancel
a blocked synchronous COM call; do not use global process termination to recover.

## F. End-to-end approval
Select a representative sanitized business fixture, independently known results and
explicitly authorized copy. Test the actual user journey, outputs, files, error recovery
and owner acceptance. Record every case using assets/test-report.md. Any unexecuted
stage remains NOT_RUN even when earlier layers passed.
