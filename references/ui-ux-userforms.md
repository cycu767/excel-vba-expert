# Desktop UserForm UI/UX audit

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Preserve behavior while improving layout
[ENG] Treat layout modernization and business-logic refactoring as separate changes. Record the baseline form size, captions, control names, containers, TabIndex, TabStop, default/cancel buttons and event handlers. Do not rename a control simply to improve its visible caption. A control's Name is an event/API dependency; its Caption is presentation.

Use alignment, consistent spacing, grouped fields and a clear primary action. Avoid a fixed-size wall of buttons; group by task and frequency. Do not apply web CSS, React layout, HTML accessibility attributes or responsive breakpoints as though they were MSForms properties. A requirement for rich searchable grids or a dependency graph may justify a worksheet or a separate HTML view, but that is an architecture decision with deployment/security consequences.

## Keyboard and validation
[DOC: MS-TABINDEX] Tab order starts at zero. [ENG] Audit it within each relevant container and through nested Frames/Pages. Test Tab, Shift+Tab, access keys, Enter and Escape. Do not make Delete/Commit an accidental Enter default. Ensure Cancel is available when validation fails; a focus trap must not prevent leaving the dialog.

Provide visible field labels, required-field cues and explanatory errors. Do not rely only on red/green color. Prefer a message that names the problem, the affected field/record and what the user can do; separate technical diagnostics into a detail view/log. Preserve valid input after an error. Wrap long text or provide a scrollable read-only details area instead of clipping it or shrinking everything.

## DPI, readability and resources
[ENG] Record the test matrix: Office build/bitness, monitor arrangement, scaling 100/125/150/200 percent, font availability and remote-desktop usage. Treat those scale values as test targets, not universal supported presets. Use actual screenshots at target scales to check clipping, overlap, focus rectangle, disabled-state contrast and minimum usable window size. Do not promise automatic per-monitor DPI correctness from a code-only change.

Preserve .frm/.frx pairing for images and serialized properties. Do not embed a third-party OCX or browser control solely for visual polish without confirming installation, bitness, licensing and organizational policy. Avoid importing fonts or binaries as hidden dependencies.

## Long-running operations
[DOC: MS-DOEVENTS] Yielding processes events. [ENG] Disable or guard duplicate actions, show progress at meaningful intervals and implement cancellation only at safe boundaries. Keep state ownership coherent if yielding. A modeless form is not a worker thread. Do not block the UI with repeated Sleep and call it asynchronous execution.

## Audit deliverable
[ENG] Produce: control/location; observed issue; severity and user effect; screenshot/code evidence; proposed visual change; event or data risk; required UI test. Mark all changes derived from screenshots alone as proposals. Retain an explicit list of behavior that must remain unchanged, including Enter/Escape/X semantics and selected-item identity.

## Source register

- [MS-TABINDEX: TabIndex property](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tabindex-property)
- [MS-CONTROLS: Controls collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/controls-collection-microsoft-forms)
- [MS-SHOW: Show method](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/show-method)
- [MS-DOEVENTS: DoEvents function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doevents-function)
- [MS-QUERYCLOSE: QueryClose event](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/queryclose-event)
