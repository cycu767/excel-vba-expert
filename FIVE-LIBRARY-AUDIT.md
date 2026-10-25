# Five-library source coverage audit (local candidate, not an upstream release)

**Scope:** verify the user's named examples against the installed, supplied
`MicrosoftDocs/VBA-Docs` snapshot using `api_preflight.py`, exact paths and
SHA-256-checked local source reads. This is not a semantic audit of every API,
Excel compilation or runtime execution. No additional Microsoft files were
needed for these examples; the original knowledge ZIP remains unchanged.

| Library | Requested examples | Local source documentation | Conclusion |
|---|---|---|---|
| VBA | ReDim, Variant, DateAdd, Collection | `Language/Reference/User-Interface-Help/{redim-statement,variant-data-type,dateadd-function,collection-object}.md` | All four source pages found |
| Excel | Workbook, Range, Chart | `api/Excel.Workbook.md`, `api/Excel.Range(object).md`, `api/Excel.Chart(object).md` | All three native Excel pages found; Graph is separate |
| MSForms | UserForm, TextBox, ListBox, ComboBox | `Language/Reference/User-Interface-Help/{userform-object,textbox-control,listbox-control,combobox-control}.md` | All four pages found; not every member is documented on the object page |
| Office | CommandBars, FileDialog | `api/Office.CommandBars.md`, `api/Office.FileDialog.md` | Both pages found |
| VBIDE | VBProject, VBComponent | `Language/Reference/Visual-Basic-Add-in-Model/objects-visual-basic-add-in-model.md` under headings `VBProject`, `VBComponent` | Both described in one grouped source page, not individual VBIDE API files |

Additional grouped VBIDE source chapters:
`collections-visual-basic-add-in-model.md`,
`events-visual-basic-add-in-model.md`,
`methods-visual-basic-add-in-model.md`,
`properties-visual-basic-add-in-model.md`.

**Actual gaps identified:** the original skill did not have deterministic
namespace-aware API preflight; search defaults returned only 8 entries and
could be misread as exhaustive; the VBIDE grouped-chapter layout was not
explicitly routed. The new helper and skill instructions close those process
gaps. They do not assert complete semantic coverage of VBA/MSForms/VBIDE.

**Known negative tests:** `MSForms.ListBox.ItemHeight` and
`MSForms.UserForm.ScaleMode` must never be approved on the basis of a ListBox
or UserForm object document. Both are recorded as unconfirmed member ownership,
not as proof of universal absence.

**Unchanged:** the external documentation archive, original notices,
imported-source hashes, knowledge index and the published v2.0.0 release.
