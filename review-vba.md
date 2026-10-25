# Review source without mistaking heuristics for proof

All steps below are engineering workflow rules, not claims that tests have already run.
## Establish review scope
List files actually read, missing binaries/designer resources, workbook context,
references, public entry points and exclusions. Run inspect_vba.py read-only with
explicit encoding where needed. A scanner candidate is not a confirmed defect.

## Review passes
1. Data integrity: wrong workbook/sheet qualification, overwrite/save/delete,
   partial output, formulas replaced by values, range size and empty data.
2. Control flow: errors swallowed, handler still active during cleanup, End,
   cancellation without cleanup, unbounded loops/retries and re-entrant events.
3. Types: implicit Variant, multidimensional arrays, ByRef mismatch, empty/error
   cells, IIf evaluating both branches and locale-dependent conversions.
4. Lifetime: child COM references, ownership of Excel instances, default UserForm
   instances, dynamic event sinks, scheduled callbacks and global state.
5. Compatibility: VBA7 declarations, pointer-sized types versus fixed DWORD,
   references/OCX dependencies and version-specific formula APIs.
6. Performance: repeated COM/cell crossings, full-column loops, repeated scans,
   volatile dependencies and unbounded logging. Measure before rewriting.
7. Maintainability: duplicated business logic, unclear side effects, hidden inputs,
   test seams and stable public contracts. Do not enforce aesthetics as defects.

## Finding format
Use ID, severity, confidence (confirmed/suspected), path/procedure/line, observed
code or runtime evidence, impact, proposed minimal patch, risk of patch and exact
regression test. Give severity an impact rationale: potential production overwrite
is different from a naming inconsistency. Report a blocked review area explicitly.

## Decision gate
Never delete an apparently unused procedure from a text-only search. Inspect
Application.Run, OnAction, ribbon callbacks, events, names, formulas, forms and
external COM invocation. Separate the review report from authorized edits.
Conclude with a prioritized small change set and actual evidence level; do not
summarize all code as safe because the heuristic scanner found nothing.
