# Audit a classic desktop UserForm

All steps below are engineering workflow rules, not claims that tests have already run.
## Gather and separate evidence
Read .frm source plus .frx presence, code-behind, handler classes and entry-point
owner. Obtain screenshots at target DPI and a list of actual user journeys where
possible. Distinguish visible layout observations from behavior inferred from code.
Do not pretend binary .frx properties were inspected by a text scanner.

## Inventory and walk-through
Create a control table: name, type, container/page, purpose, TabIndex, default/cancel
behavior, input restrictions, event handlers and model field. Verify duplicate
names/accelerators, no-selection states and whether display text is mistaken for
persistent key. MultiPage and TabStrip have different containment contracts.

Trace Initialize -> caller context injection -> Show -> Activate -> edits ->
validation -> commit/cancel -> Hide/Unload -> owner cleanup. Note that activation
can repeat. For modeless forms, verify the owner retains the instance and reacts to
workbook close. For runtime controls, verify persistent WithEvents sinks.

## UX tests specific to Windows desktop
Traverse with Tab/Shift+Tab, Enter, Escape, accelerators and Alt+F4. Check focus after
validation failure, readable wrapped messages, visible mandatory-field meaning and
an unambiguous Cancel. Test common target scaling values (for example 100/125/150/200
percent) as a test matrix, not a claim of automatic DPI support. Check contrast,
labels, clipped controls, scrolling and layout inside Frames and Pages.

## Behavior-preserving correction
Prioritize blocked actions and data loss above cosmetic issues. Keep existing
control names and event signatures unless a migration is authorized. Change layout
separately from business rules. Do not offer CSS/React properties as MSForms APIs.
When UI demand exceeds Forms, compare a worksheet view or explicit external HTML
view, including security/offline/deployment costs; do not choose one by fashion.

## Deliverable and gates
Report control-level evidence, severity, screenshot/code location, proposed change
and test journey. Deliver code/designer changes only in an authorized copy, keep
.frx companions, and re-import on real Excel before declaring the form deployable.
Source-only audit cannot certify screen-reader behavior, DPI rendering or OCX load.
