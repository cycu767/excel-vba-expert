# Code assets: original engineering examples

Status: reviewed as source text; NOT compiled or executed in Excel in this build.
Use a disposable workbook and explicit permission. These files are not installers.

| Asset | Placement | Dependency / limitation |
|---|---|---|
| MApplicationState.bas | Import as standard module | Excel object library; value-copy overwrite needs authorization; no rollback |
| WorksheetChange.form-code.txt | Paste in target worksheet module | Not a UserForm despite the generic snippet suffix; modifies B2:B1000 literal text |
| CButtonHandler.cls | Import as class module | MSForms reference via real UserForm; retain sink instances |
| DynamicButtons.form-code.txt | Paste in empty real UserForm | Needs CButtonHandler; owner must unload/dispose; no invented designer/frx |
| MCharts.bas | Import standard module | Two-column data on same sheet; named embedded chart only |
| MWinApi.bas | Import standard module | Windows; compile separately on 32/64-bit; HWND vs DWORD |
| MValidation.bas | Import standard module | Pure quantity parser, not a general currency parser |
| MValidationTests.bas | Import standard module | Requires MValidation; manual runner raises on failure |

Import .bas/.cls files through VBE File > Import File. Do not paste exported
Attribute/VERSION headers into a code pane. Text snippets have no such headers.
Do not paste a second Option Explicit into an existing module. Merge with existing
events; a worksheet cannot contain two Worksheet_Change procedures.
All assets use ASCII to avoid silent legacy VBE codepage conversion. Preserve the
encoding of existing exports. Existing .frm and .frx must travel together.

Review the preconditions before reusing templates. Passing these examples through
Python syntax checks does not validate VBA. The real acceptance procedure is in
[Windows acceptance](../tests/WINDOWS-ACCEPTANCE.md).
