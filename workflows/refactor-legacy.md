# Modernize legacy VBA while preserving behavior

All steps below are engineering workflow rules, not claims that tests have already run.
## Baseline before surgery
Inventory modules, .frm/.frx pairs, add-ins, references, file encodings, public entry
points, hidden sheets/names, external links and workbook opening/closing behavior.
Use a copy and version-control exported sources. Record stable input fixtures,
expected reports, calculation totals, rounding, form paths and run times. A business
owner must resolve undocumented ambiguous behavior before it is changed.

## Dependency map
Trace explicit calls plus indirect entry points: workbook/worksheet events,
Application.Run string names, shapes OnAction, Ribbon XML callbacks, UDFs used in
cells, scheduled procedures, form handlers and external scripts. A procedure with
zero text references may still be live. Mark dynamic edges unresolved rather than
removing their targets. Record resource and application-state ownership.

## Refactoring order
1. Add characterization tests and diagnostic boundaries without changing results.
2. Qualify object references locally, testing with a different workbook active.
3. Introduce explicit arguments around one business operation at a time.
4. Separate validation/computation from UI and writes; preserve public wrappers.
5. Replace broad suppression with narrow, documented error paths and state cleanup.
6. Repair one compatibility boundary (for example a Windows API declaration) with
   its own tests, not a global Long-to-LongPtr replacement.
7. Optimize only measured hotspots after behavioral equivalence is established.

## Stop and recovery conditions
Stop on changed totals, file format loss, dropped references, missing frx assets,
signature invalidation, unapproved new dependency or an unrepeatable baseline.
Restore the known-good copy; do not repeatedly patch production to regain a result.
Splitting a form or module must not silently change event order or instance lifetime.

## Deliverable
Provide baseline matrix, dependency uncertainties, staged patches, compatibility
notes, migration/import sequence, rollback plan and regression evidence per stage.
Distinguish intentional business changes from behavior-preserving refactoring.
No static tool here proves absence of dead code or complete call-graph coverage.
