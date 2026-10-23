# Develop a feature without changing business contracts

All steps below are engineering workflow rules, not claims that tests have already run.
## Inputs and preflight
Read the actual modules and workbook layout, not just a screenshot or names. Identify
Excel version/build, Office bitness (not Windows bitness), file format, locale,
references, entry point, workbook ownership and requested output. Mark unknowns.
Establish an authorized working copy and baseline results before changing code.
State non-goals: do not silently replace calculation rules, external integrations,
form instance ownership or data sources while implementing a visual feature.

## Design and implementation
1. Write acceptance examples with normal, empty, invalid and boundary inputs. Map
   inputs -> pure validation/transformation -> explicit worksheet/form outputs.
2. Decide whether a change belongs in a standard module, class, sheet event or
   UserForm code-behind. Keep existing exported names when callers depend on them.
3. Resolve every uncertain member against the local index and the actual documented
   signature. Ask for permission before online lookup; do not guess a method.
4. Pass Workbook/Worksheet/Range explicitly across boundaries. Use ThisWorkbook
   only when the code-containing workbook is truly the target; add-ins are a
   counterexample. Keep object ownership separate from the active UI context.
5. Validate the full input and target permissions before the first write. Define
   partial-write recovery; a VBA error handler is not a transaction.
6. Implement the smallest cohesive patch. Capture exact application state before
   mutation and preserve primary errors through cleanup. Do not insert global
   On Error Resume Next or use End as an emergency return.
7. Record changed procedures, caller effects, dependencies, import encoding and
   any .frm/.frx pair implications. Keep a patch or version-control diff.

## Gates and delivery
Run available local static tools; compile in VBE only when Excel is really present.
Then test normal behavior, invalid input, user cancellation, already-disabled events,
another workbook activated, repeated execution and injected error. For forms, add
keyboard/close/lifecycle tests. For files, verify names, format, ownership and contents
before promoting a generated output. Reconcile numerical outputs against baseline
using the business-defined rounding rule, not an invented tolerance.

Deliver the patch or complete small module, placement instructions, source IDs,
expected behavior and an evidence table: STATIC / COMPILED / RUNTIME / UI /
END_TO_END with actual PASS/FAIL/NOT_RUN. Do not claim the feature is production
ready if only source review was possible.
