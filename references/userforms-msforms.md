# UserForms and Microsoft Forms: lifecycle and control contracts

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Establish the form contract
[ENG] Inventory .frm AND its .frx companion, form class name, exported attributes, references/OCX dependencies, controls, containers, event handlers, entry points, result outputs and workbook targets. A screenshot alone cannot establish event wiring or hidden controls. A textual .frm may not expose every property serialized in its binary .frx. Never invent a replacement .frx or claim a generated text file is a tested exported form.

## Instance lifecycle
[DOC: MS-USERFORM, MS-INITIALIZE, MS-ACTIVATE, MS-HIDE, MS-SHOW, MS-UNLOAD] Initialization, activation, showing, hiding and unloading are distinct. [ENG] Prefer an explicit `Dim dlg As FEditRecord: Set dlg = New FEditRecord` when ownership matters. Inject validated context through an initialization method; do not assume Initialize already has caller-supplied data. Reserve Activate for idempotent display work, not unguarded data insertion.

For a modal dialog, keep draft values local, accept/cancel by setting an explicit result and hiding, read the result in the owner, then unload and release the instance. For modeless forms, retain an owner reference after the launch procedure returns and handle workbook closure. Avoid re-accessing a default instance after Unload: it may create/load a new instance instead of returning the result you expected.

[DOC: MS-QUERYCLOSE] Distinguish X/control-menu close from code unload and application shutdown. [ENG] Handle cancel consistently but do not cancel all close causes indiscriminately. Use an explicit Dispose/cleanup path; Terminate is not a transactional finally block. [UNVERIFIED] The Terminate reference was unavailable in this build; verify detailed event ordering on the target host before relying on it.

## Control selection matrix
[ENG] These are design choices; verify each specific property/event in the installed Object Browser before adding it.

| Control | Appropriate use | Main audit check |
|---|---|---|
| CommandButton | One explicit action | Duplicate click protection; destructive action not default |
| TextBox | Editable text / validated numeric input | Empty/error handling; Change versus commit validation |
| Label | Field name/instructions | Wrapping, association and readable content |
| ComboBox | One choice, optionally editable | Whether arbitrary typed values are allowed; stable key |
| ListBox | List selection | No selection, multiselect, key versus display column |
| CheckBox | Boolean/tri-state only when intended | Null/indeterminate policy |
| OptionButton | Mutually exclusive choice | Correct Frame/container and grouping |
| Frame | Visual and logical group | Child coordinates and keyboard order |
| MultiPage | Distinct control pages | Correct Pages/Value and per-page validation |
| TabStrip | Reuse controls for different data sets | Update shared controls explicitly on selection |
| Image | Nonessential visual cue | Resource presence; no critical text only in image |
| ScrollBar / SpinButton | Bounded numeric adjustment | Boundaries, step and keyboard equivalence |

[DOC: MS-MULTIPAGE, MS-TABSTRIP] MultiPage has Pages; TabStrip has Tabs. Do not replace one with the other without redesigning data/control ownership.

## Namespace and property verification

[ENG] `MSForms.ListBox` is not .NET WinForms `ListBox` and an Access form
is not `MSForms.UserForm`. Do not assume `ListBox.ItemHeight` or
`UserForm.ScaleMode` exists because similarly named APIs exist elsewhere.
Use `api_preflight.py` for the exact class/member combination and read the
matched Microsoft source. `OWNER_TYPE_CONTEXT` proves the form/control is
documented **not** that a requested member exists. For rendering dimensions,
verify units (points versus pixels), read/write status, DPI and empirical UI
results in the target Excel version. [UNVERIFIED] A font-based row-height
estimate is a layout approximation, not guaranteed pixel-perfect.

## Event wiring and loading
[DOC: MS-CONTROLS-ADD, MS-DIM] Runtime Controls.Add uses a ProgID; WithEvents variables require a class and cannot be arrays or As New. [ENG] Retain one handler object per dynamic control in a collection, otherwise its events can disappear with the handler's lifetime. See `CButtonHandler.cls` and `DynamicButtons.form-code.txt`. Do not assume a procedure named after a new control automatically receives events.

Use a form-local `mLoading` or `mBusy` flag around programmatic changes. Reset it on every error path. On a Change event, avoid committing business transactions or showing modal warnings for every keystroke. Validate at meaningful boundaries, preserve input after failure, focus the failing field and keep the commit operation separate.

## Completion gates
[ENG] Test initial show, hide/show, repeated activation, X, Cancel, Enter, Escape, Alt+F4, invalid input, empty lists, double-click/duplicate action, workbook close and target workbook switching. Test dynamic controls after the creator returns. Record observed behavior, not inferred screenshot behavior. UI and .frx import tests require real Excel/MSForms.

## Source register

- [MS-USERFORM: UserForm object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/userform-object)
- [MS-INITIALIZE: Initialize event](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event)
- [MS-ACTIVATE: Activate and Deactivate events](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/activate-deactivate-events)
- [MS-HIDE: Hide method](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/hide-method)
- [MS-SHOW: Show method](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/show-method)
- [MS-UNLOAD: Unload statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unload-statement)
- [MS-QUERYCLOSE: QueryClose event](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/queryclose-event)
- [MS-CONTROLS-ADD: Add method (Microsoft Forms)](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms)
- [MS-CONTROLS: Controls collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/controls-collection-microsoft-forms)
- [MS-DIM: Dim statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
- [MS-MULTIPAGE: MultiPage control](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/multipage-control)
- [MS-TABSTRIP: TabStrip control](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/tabstrip-control)
