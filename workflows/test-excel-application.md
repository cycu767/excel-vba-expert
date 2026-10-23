# Plan and execute tests with honest evidence levels

All steps below are engineering workflow rules, not claims that tests have already run.
## Define the test contract
For each test record fixture, preconditions, action, oracle, expected side effects,
expected application state, cleanup and evidence level. Pure logic can be exercised
through a VBA runner, but code must actually run before calling it a unit-test pass.
A Python replica of an algorithm is not execution of that VBA code.

## Layered execution
1. Source review: validate exports, file pairs, hardcoded paths, risky state changes
   and assumptions. Record heuristic findings separately from compilation errors.
2. VBE compilation: open an approved disposable project, check references and use
   Debug > Compile. Test appropriate conditional-compilation branches on both Office
   bitnesses where supported. Do not use undocumented UI command IDs as proof.
3. Pure VBA tests: import MValidation and MValidationTests into a test workbook;
   manually run the test procedure. Assertions must raise on failure for batch use.
4. Workbook integration: test explicit workbook identity, empty/single/multi-area
   ranges, error values, formulas, protection, merged cells, names and saves.
5. State/recovery: inject errors after each mutation; enter with events already off
   and calculation manual. Assert exact restoration and preservation of primary
   failure. VBE Reset/process crash are outside ordinary finally-like guarantees.
6. UserForm: test navigation, DPI, every close path, repeated activation, dynamic
   buttons, no selection, invalid input, double action and workbook switching.
7. COM: run the opt-in blank-workbook smoke test; separately authorize any macro
   invocation or production-copy integration. Mock cleanup tests are not real COM.
8. End-to-end: use a representative user journey and reconcile the final artifact
   against the business oracle. A correct cell is not proof that export/email/save
   or downstream consumption worked.

## Reporting
Use PASS / FAIL / NOT_RUN / BLOCKED, not a single overall green indicator. State
platform, Excel build/bitness, fixture identity, timestamp and observed values.
Record failure cleanup and retained files. Do not delete a user's evidence in test
cleanup. The delivered skill's Windows acceptance matrix is an executable human
procedure; its presence alone does not mean the tests were performed.
