# Evidence levels and acceptance tests

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Use explicit evidence levels
[ENG] Report each result separately: DOCUMENTATION_REVIEW; STATIC_REVIEW; PYTHON_TEST; MOCK_TEST; VBA_COMPILE; VBA_RUNTIME; COM_RUNTIME; UI_TEST; END_TO_END. PASS in one level does not imply PASS in another. NOT_RUN, BLOCKED and FAIL are meaningful outcomes, not variants of PASS.

A Python syntax check does not check VBA syntax. A regex linter does not compile VBA. A mocked COM object does not run Excel. A successful real COM calculation does not test UserForm keyboard behavior. Reading a prompt does not establish that Qwen invoked the intended skill in VS Code.

## Test layers
[ENG] Extract pure transformation functions for small deterministic VBA tests; use explicit assertion helpers that raise errors rather than leaving an unattended process halted at Debug.Assert. Use fixtures for integrations: tables, formulas, named ranges and multiple workbooks. For regression, compare values, formulas, counts, keys, ordering, rounding and selected visual outputs according to the contract.

Test event entry and its domain procedure separately. Test form initialization, validation, cancel, close and modeless lifetime in Excel. Test file operations with new paths, collisions, denied writes, partial failure and reopen validation. Test COM cleanup on normal and exceptional paths. If shutdown cannot be observed, do not declare there are no orphan processes.

## Fault and boundary matrix
[ENG] Include zero/one/many records; duplicates; Null/Empty/error/overflow; locale-specific numbers/dates; missing sheet/table/name; formulas versus values; protected workbook; read-only file; cancelled dialog; stale external data; disabled events/manual calculation on entry; injected failure after state mutation; second workbook becoming active; repeated run; 32/64-bit declarations and deployed controls.

Map every fixed bug to a reproducer and at least one neighboring case. Record expected output BEFORE the fix so that the test cannot simply approve the new wrong result. For financial logic, use approved invariants and reconciliation totals rather than an invented tolerance.

## The bundled checks
[ENG] `validate_skill.py` validates this package's structure, links, canonical source index and routing scenarios. `python -m unittest discover -s tests -v` exercises search, local indexing, static inspection, ZIP safety and mocked lifecycle logic. `tests/scenarios.json` is a prompt/rubric set; automated routing checks are not a behavioral evaluation of Copilot or Qwen. `tests/WINDOWS-ACCEPTANCE.md` contains the remaining live acceptance steps.

Run `excel_smoke_test.py` without an execution flag for dependency detection. Use --run-local-excel only on an authorized Windows test host. Record Office build/bitness and results. Its blank-workbook test does not compile the VBA templates or validate macro security for arbitrary files.

## Report format
[ENG] For each test: ID, level, environment, fixture/hash, action, expected, actual, status, evidence location, remaining risk. State blockers plainly and deliver the next executable test steps, not a claim that testing is complete because the files exist.

## Source register

- [MS-ASSERT: Debug.Assert](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assert-method)
- [MS-DEBUG-MENU: Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)
- [MS-CHANGE: Worksheet.Change](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.change)
- [MS-EVENTS: Application.EnableEvents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents)
- [MS-64BIT: 64-bit VBA overview](https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview)
