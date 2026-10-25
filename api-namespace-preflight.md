# API ownership and source preflight (VBA / Excel / MSForms / Office / VBIDE)

Evidence types: **DOC** means verified source text; **ENG** is a safety rule;
**UNVERIFIED** means host-specific or unsupported by retrieved source. The source
snapshot is not a live type-library dump or Excel compile/runtime confirmation.

## Before writing a new property, method, event or constant

1. Identify the *exact owning type* and library. Distinguish `Excel.Chart`
   from legacy Graph, `MSForms.ListBox` from .NET WinForms and Access controls,
   and `Office.CommandBars` from VBE command bars. Treat `Scripting.Dictionary`
   as an external Microsoft Scripting Runtime object, not a native `Collection`.
2. Run `scripts/api_preflight.py --namespace <library> --symbol <Object.Member>`.
   Its evidence is path-scoped and SHA-256-verified. **A container's documentation
   alone never proves that an arbitrary member belongs to that type.**
3. Open the returned source with `offline_docs.py read <source_path> --expand-includes`.
   For grouped VBIDE chapters, read the indicated original line range. Confirm
   the member's owning type, host/version, read/write status, units and parameters.
4. If the direct source is incomplete, search again with `offline_docs.py search
   "<term>" --scope all --limit 60`, exact source-title variants and explicit
   source paths. Never infer absence from default 8 results, category labels or
   one `--exact` miss. Use the Object Browser in the target Excel/VBE to resolve
   ambiguity when the local documents group members or omit applicability.
5. Do not implement an unconfirmed member as if it were supported. Mark
   `NOT_CONFIRMED_IN_SNAPSHOT` / `MEMBER_NOT_CONFIRMED_OWNER_ONLY` / `UNVERIFIED`,
   ask for targeted host check where necessary, or choose a documented alternative.
   Do not invent compatibility shims for non-existent members.
6. After a verified contract, implement in a **working copy** and compile/run in
   Excel only if authorized and available. Record `DOC`, `STATIC_REVIEW`,
   `VBA_COMPILED`, `VBA_RUNTIME` separately, not as interchangeable PASS values.

## Command examples

PowerShell (from any current directory):

```powershell
$skillRoot = Join-Path $HOME '.copilot\skills\excel-vba-expert'
$preflight = Join-Path $skillRoot 'scripts\api_preflight.py'
$docs = Join-Path $skillRoot 'scripts\offline_docs.py'
Test-Path -LiteralPath $preflight
python $preflight --namespace msforms --symbol ListBox.ItemHeight --json
python $preflight --namespace msforms --symbol ComboBox --json
python $preflight --namespace excel --symbol Chart --json
python $preflight --namespace vbide --symbol VBComponent --json
python $docs search 'ComboBox control' --exact --scope all --limit 60
python $docs read 'Language/Reference/User-Interface-Help/combobox-control.md' --expand-includes
```

Do **not** use `%USERPROFILE%` in PowerShell or `$env:USERPROFILE` in CMD.
A PowerShell profile-loading warning is separate from running these Python tools;
do not weaken corporate execution policies to work around it.

## What each namespace means in this snapshot

| Namespace | Source ownership | Do not confuse with |
|---|---|---|
| VBA | `Language/Reference/User-Interface-Help/*` language pages | Windows API / Scripting Runtime / MSForms merely because they use VBA |
| Excel | `api/Excel.*` without legacy `-graph` pages | Graph Chart, Access Range, Outlook objects |
| MSForms | `Language/Reference/User-Interface-Help/*-control.md`, `userform-object.md` and associated member pages | Access controls, VB6, .NET WinForms |
| Office | `api/Office.*` | `Excel.*` application-specific objects or VBE UI |
| VBIDE | `Language/Reference/Visual-Basic-Add-in-Model/{objects,collections,properties,methods,events}-visual-basic-add-in-model.md` | Excel Workbook's `VBProject` *accessor* alone |

For VBIDE, original documents group multiple objects and members under headings;
the lack of `api/VBIDE.VBComponent.md` is **not** evidence the VBComponent
object is missing. Check the specific heading and owner, then verify host API
through the Object Browser. Access to the VBA project object model may be
blocked by policy: **never enable Trust access automatically**.

## Known counterexamples and limitations

- `MSForms.ListBox.ItemHeight` is **not** established by the MSForms ListBox
  source. .NET WinForms has an `ItemHeight` API; do not transfer it to MSForms.
  For layout, use documented MSForms dimensions and test rendered behavior.
- `MSForms.UserForm.ScaleMode` is **not** established by the UserForm source.
  An identically named member in Access does not establish MSForms support.
- `Excel.Chart` uses the native Excel Chart model; legacy Microsoft Graph is a
  different library. Source path and exact owning type must both agree.
- A member page may describe a property shared by multiple controls but not all
  of them. The preflight labels generic member matches as **candidates only**.
- A database search miss proves only that the *supplied snapshot* did not
  confirm the symbol under the searched routes. It does not prove non-existence.

## Status vocabulary (`api_preflight.py`)

- `SOURCE_DOCUMENT_FOUND`: direct object/standard language document or VBIDE
  section found; read before use.
- `MEMBER_DOCUMENT_FOUND_CHECK_OWNER`: member document found but owner and
  contract still need confirmation in the source and target host.
- `CANDIDATES_REQUIRE_MEMBER_VERIFICATION`: related grouped/generic source
  exists; direct member ownership remains unresolved.
- `MEMBER_NOT_CONFIRMED_OWNER_ONLY`: owning object page exists, but does **not**
  establish requested member; never treat as supported without further evidence.
- `NOT_CONFIRMED_IN_SNAPSHOT`: no matched source found in that namespace; expand
  manual investigation. Neither outcome proves API non-existence.
