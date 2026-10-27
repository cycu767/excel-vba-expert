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

`compare_workbook_values.py` and `compare_vba_source.py` need no Excel, no network
and no third-party package, and refuse to report equality on an empty or
truncated read. They compare cached cell values and decompressed VBA source
respectively; neither is a malware verdict, a compile, or a behavioural test.

## Verify the primitive before building a check on it
[ENG] Before an integrity control is trusted, test the primitive it rests on
against an independent implementation. A hashing routine that always fails and
returns an empty string makes every comparison `"" <> ""` and therefore declares
every artefact, including a swapped one, intact. The control is not merely
unreliable; it is fictitious, and it will look like the strongest evidence in the
report.

[ENG] A compile or load probe must actually exercise the paths it claims to
cover. A language that compiles lazily along the call graph will accept a module
containing an error in a procedure the probe never reaches, so the probe reports
success and the defect reaches the user. Prove coverage by invoking the real
entry points, and do not wrap such a probe in error handling that would swallow
the very compile error it exists to detect.

## Proofs that cannot pass vacuously
[ENG] Every equality, coverage or "no findings" check must fail when it reads
nothing. Assert the size of what was read before comparing it: component count,
total lines, row count or cell count above a floor taken from the fixture. An
intersection of two empty sets is empty, a loop over zero items reports no
differences, and a diff of two truncated reads reports success. Such a result is
a tool failure reported as a pass, which is worse than a failure.

[ENG] Make the negative case part of the test. Feed the checker one deliberate
difference and require it to report that difference, by name and location. A check
that has only ever returned "identical" has not been demonstrated to detect
anything.

[ENG] Guard against silent truncation. If a reader caps rows, bytes or records,
raise rather than return a partial result, and make the cap an explicit option the
caller can raise. A comparison over a truncated subset must not be labelled
IDENTICAL.

[ENG] An assertion that does not fail loudly on missing or unparsed data is worse
than no assertion. A numeric parse that yields `None`/Empty and then flows into a
comparison produces either a spurious failure or a vacuous pass, and the reader
cannot tell which. Validate that parsing produced a value before asserting on it,
and report the parse failure as its own distinct outcome.

[ENG] Prove a guard by the state of what it protects, not by the text of its
refusal. Capture the protected data before and after the attempted operation and
require it to be unchanged. Pair every refusal test with a positive test showing
the permitted operation still succeeds: a guard that always refuses must not be
able to earn a green result.

[ENG] A failing gate is evidence about the test before it is evidence about the
product. Before diagnosing a product defect, confirm the test's own inputs exist
in the fixture: a search term matching zero records produces an empty list and a
correct refusal that looks exactly like a broken control. Mass failure across
otherwise unrelated gates points at the environment, not at a product regression.

[ENG] State the scope of every check in its own output, not only in the
documentation: which layers, sheets, components or properties were compared and
which were not. A green result whose scope nobody can read is not evidence.

## Verification effort proportional to risk
[ENG] Match the depth of verification to the consequence of being wrong, and say
so explicitly. A change to a configuration table in one month may need value
comparison plus a rule-conflict check. A change to calculation code, to many
historical periods, or to a package handed to someone else needs the original
preserved, source-level comparison, value comparison, and a check that unrelated
sheets and states are untouched. Do not spend the expensive checks on trivia and
do not substitute a cheap check when the expensive one is what the risk requires.

[ENG] A verification script that aborts on a precondition it cannot satisfy has
not verified the work. If a gate exits before the substantive checks run, the run
proves only that the gate fired. Either make the gate valid for the current
situation or split the script so the checks that can run do run, and report which
ones did. Never report the exit code of an aborted script as a pass.

## Derived sheets are snapshots, not live views
[ENG] A validation or diagnostics sheet written by a procedure shows the state at
its last run, not the current state. Run the procedure that writes it, then read
it. Reading a stale control sheet and reporting its contents as the present
condition is a false result produced by a real file. If the writer needs a quiet
mode to avoid a modal prompt, enable it first; an unattended run blocked on an
invisible dialog is indistinguishable from a hang.

## Proving a rewrite is equivalent
[ENG] Prove equivalence of a refactor or migration by running the old code and the
new code over the same complete dataset and comparing values, not structure. Create
the reference output *before* the change, archive it, and compare the new run
against it: row count, per-row values, totals, exception counts, and the closure
check. Define the permitted differences from the approved change contract (which
may include a timestamp); any other difference requires explicit investigation.

[ENG] If the business contract declares prior periods immutable, treat historical
rows generated under earlier rules as snapshots. Do not silently recalculate them
using today's configuration. Any approved restatement or correction needs its own
signed-off migration plan, original preserved, per-cell change map, reconciliation
and post-edit proof of the exact approved delta. Other projects may legitimately
require historical restatement: determine the contract first rather than applying
a universal ban on historical recalculation or re-import.

[ENG] A gate becomes credible only after it has been run against code known to be
wrong and failed. Report how many checks fell and which. A gate that has only ever
passed has not been shown to measure the effect; it may be matching text in the
source instead of the state of the data.

[ENG] An independent replica used as a reference must implement the same precedence
rules as the production code, and must be compared over the full history, not one
period. A replica that resolves ties differently reports thousands of false
disagreements, and the tempting "fix" is then to change production to match the
replica. Build the replica from the production rules, and treat a simulator as a
before/after instrument, not a source of absolute values.

[ENG] Every corrective edit to data needs an explicit double-count check: confirm
the document appears once in the source files and once in any manual-assignment
table, and confirm that exclusion tables cannot add the amount a second time. A
"first key wins" rule over duplicate keys is acceptable only while duplicates are
detected and reported, never while they are silently tolerated.

## Report format
[ENG] For each test: ID, level, environment, fixture/hash, action, expected, actual, status, evidence location, remaining risk. State blockers plainly and deliver the next executable test steps, not a claim that testing is complete because the files exist.

## Source register

- [MS-ASSERT: Debug.Assert](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assert-method)
- [MS-DEBUG-MENU: Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)
- [MS-CHANGE: Worksheet.Change](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.change)
- [MS-EVENTS: Application.EnableEvents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents)
- [MS-64BIT: 64-bit VBA overview](https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview)
