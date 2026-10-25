# Excel VBA Expert v2.0.1 — AI Agent Skill with Offline Microsoft VBA Documentation

**An offline-first AI Agent Skill for professional Microsoft Excel VBA development, debugging, code review, maintenance, and automation.**

**Release date:** October 23, 2026  
**Project status:** An extension of the existing Excel VBA Expert Skill, not a new project.

Excel VBA Expert extends compatible AI coding agents with specialized engineering instructions, 19 expert modules, 7 development workflows, and a searchable offline knowledge base containing 38,187 Microsoft VBA documentation files.

The skill is designed for professional VBA development, including complex, business-critical, and long-lived Excel applications.

**This is not a standalone AI model, Excel add-in, or documentation viewer.** It is an installable AI Agent Skill that guides AI assistants through VBA engineering tasks and provides local tools for consulting official Microsoft documentation.

## How It Works

AI Coding Agent → Excel VBA Expert Skill → Engineering Workflows + Offline Microsoft Documentation → Source-backed VBA Assistance

The project is distributed in two packages:

| Package | Contents |
|---|---|
| `skill.zip` | Core AI Agent Skill, engineering instructions, expert modules, workflows, scripts, examples, and tests |
| `excel-vba-knowledge.zip` | Complete offline Microsoft VBA documentation snapshot, full-text search index, portable indexes, and source metadata |

**Both packages are required for full offline functionality.**

The knowledge base is distributed separately through GitHub Releases to keep the main repository and core skill lightweight.

## Two Packages, One Installation

`skill.zip` contains the skill, all 19 preserved expert modules, 7 workflows, existing tools and examples, and the new utilities for accessing the complete offline documentation.

`excel-vba-knowledge.zip` contains the **actual Microsoft VBA-Docs source files**: all 38,187 Markdown documents from the supplied repository snapshot, the remaining repository files, original licenses, a full-text search index, and portable JSONL indexes.

There is no need to download the Microsoft repository separately or rebuild the search index.

After extracting **both archives into the same directory**, the expected structure is:

```text
%USERPROFILE%\.copilot\skills\
  excel-vba-expert\
    SKILL.md
    scripts\offline_docs.py
    metadata\knowledge-config.json
  excel-vba-expert-knowledge\
    manifest.json
    index.sqlite
    documents.jsonl
    repository\
      api\Excel.Range.Find.md
      Language\...
      LICENSE
      LICENSE-CODE
      ThirdPartyNotices
```

The knowledge directory is a sibling of the skill directory, **not a second skill**. It does not contain a `SKILL.md` file.

Both ZIP archives already contain their own root directories. Do not create an additional nesting level when extracting them.

The extracted files contain approximately 177 MB of logical data. Actual disk usage may be higher because of the large number of small files. Reserve at least **400 MB of free disk space**.

## Windows Installation and Updates

Keep both downloaded ZIP archives in the same directory.

The following PowerShell script moves existing installation directories to a backup location outside the Skills discovery directory before installing the new version. It does not delete the previous installation.

Open PowerShell in the directory containing both ZIP files and execute:

```powershell
$ErrorActionPreference = "Stop"

foreach ($zip in @("skill.zip", "excel-vba-knowledge.zip")) {
    if (-not (Test-Path -LiteralPath $zip -PathType Leaf)) {
        throw "Missing archive: $zip"
    }
}

$base = Join-Path $HOME ".copilot\skills"
$backup = Join-Path $HOME (".copilot\skill-backups\vba-" + (Get-Date -Format "yyyyMMdd-HHmmss"))

New-Item -ItemType Directory -Force -Path $base | Out-Null

foreach ($name in @("excel-vba-expert", "excel-vba-expert-knowledge")) {
    $existing = Join-Path $base $name

    if (Test-Path -LiteralPath $existing) {
        New-Item -ItemType Directory -Force -Path $backup | Out-Null
        Move-Item -LiteralPath $existing -Destination (Join-Path $backup $name)
    }
}

Expand-Archive -LiteralPath .\skill.zip -DestinationPath $base
Expand-Archive -LiteralPath .\excel-vba-knowledge.zip -DestinationPath $base

$skill = Join-Path $base "excel-vba-expert"

py -3 "$skill\scripts\offline_docs.py" verify

if ($LASTEXITCODE -ne 0) {
    throw "Knowledge verification failed; keep the backup."
}
```

Do not delete the backup of the previous version until the new installation has been verified.

After installation, start a new conversation with your AI coding agent and select or explicitly reference `excel-vba-expert`.

Skill discovery depends on the AI client, extension, and configuration. Installation in one agent does not automatically make the skill available to every other agent.

## Requirements

- Python 3.10 or later.
- Python's built-in `sqlite3` module with SQLite FTS5 support.
- Sufficient disk space for the extracted documentation corpus.

The following are **not required** for offline documentation search:

- Additional pip packages.
- Git.
- A vector database server.
- An embedding model.
- External API keys.
- Microsoft Excel.
- An internet connection.

The original Markdown files can also be read directly without Python.

To verify SQLite FTS5 support without internet access:

```powershell
py -3 -c "import sqlite3; c=sqlite3.connect(':memory:'); c.execute('CREATE VIRTUAL TABLE t USING fts5(x)'); print('FTS5 OK')"
```

## Offline Documentation Verification

The following commands can be executed from any working directory.

The knowledge base is located relative to the installed skill, not relative to the current project:

```powershell
$skill = Join-Path $HOME ".copilot\skills\excel-vba-expert"

py -3 "$skill\scripts\offline_docs.py" search "Range.Find" --exact --json
py -3 "$skill\scripts\offline_docs.py" read "Range.Find" --expand-includes
py -3 "$skill\scripts\offline_docs.py" read "UserForm.Show"
py -3 "$skill\scripts\offline_docs.py" read "xlWhole"
py -3 "$skill\scripts\offline_docs.py" links "Range.Find"
```

By default, the `read` command returns the **complete source document**, not an AI-generated summary.

Every document read includes SHA-256 file integrity verification.

Additional options include:

- `--section Remarks` — retrieve a specific section.
- `--start-line` — specify the first source line.
- `--end-line` — specify the last source line.
- `--json` — return structured JSON output.

When a name is ambiguous, use the exact document path returned by the search command.

For example:

```powershell
py -3 "$skill\scripts\offline_docs.py" read "api/Excel.Range(object).md"
```

This distinguishes Excel's `Range` object from similarly named objects in other Office applications.

To search documentation covering other Office products:

```powershell
py -3 "$skill\scripts\offline_docs.py" search "Document" --scope all
```

## Example AI Agent Verification Prompt

Use the following prompt to verify that an AI coding agent can discover the skill and retrieve local documentation:

> Use the `excel-vba-expert` skill.
>
> Without accessing the internet, search the local documentation for `Range.Find`.
>
> Read the complete source document and identify the local file path and source lines containing the method signature, parameters, and remarks.
>
> Then propose safe VBA code based on the retrieved Microsoft documentation and the relevant expert engineering module.
>
> Do not launch Microsoft Excel or modify any project files.
>
> Clearly distinguish documented API behavior from engineering recommendations.

The same verification approach can be applied to:

- `Workbooks.Open`
- `Worksheet.Change`
- `UserForm.Show`
- `ChartObjects.Add`

If a section is absent from the original Microsoft documentation, the agent must explicitly report that absence rather than invent missing content.

See [Offline Verification Results](OFFLINE-VERIFICATION.md) for the recorded results of these five documentation retrieval tests.

## Real-World AI Agent Integration Test

The skill has been tested with an AI coding agent running in **Visual Studio Code**.

In a practical `Excel.Range.Find` documentation retrieval test, the agent:

1. Discovered and loaded the `excel-vba-expert` skill.
2. Invoked the local `offline_docs.py` utility.
3. Retrieved Microsoft documentation for `Excel.Range.Find`.
4. Identified the method signature and parameters.
5. Correctly identified the four persistent search settings: `LookIn`, `LookAt`, `SearchOrder`, and `MatchByte`.
6. Reported the source file path, metadata, and integrity verification information.

The agent's execution history showed actual local tool invocations, providing evidence of successful skill discovery and documentation retrieval.

**Test scope:** This validates the AI agent's ability to use the local skill and documentation tools. It does not constitute end-to-end validation of VBA compilation, Excel automation, or execution of generated code inside Microsoft Excel.

## Updating the Knowledge Base from a New Local Repository Snapshot

Do not overwrite the existing engineering rules or a working knowledge base during an update.

Import the new documentation snapshot into a **new destination directory**:

```powershell
py -3 "$skill\scripts\import_knowledge.py" --source ".\VBA-Docs-main.zip" --dest "D:\VBA-KB-next"

py -3 "$skill\scripts\offline_docs.py" --kb "D:\VBA-KB-next" verify

py -3 "$skill\scripts\offline_docs.py" --kb "D:\VBA-KB-next" read "Range.Find"
```

The `--source` argument can point to either a ZIP archive or an extracted repository directory.

The import process:

- Does not download external content.
- Does not execute content from the imported repository.
- Does not overwrite an existing destination directory.
- Excludes only `.git` metadata when importing from a Git working directory.
- Preserves source documentation and license files without modification.

After verifying the new knowledge base, replace the existing sibling knowledge directory while retaining a backup.

Alternatively, set the `EXCEL_VBA_KB` environment variable in the agent process to point to the new knowledge directory.

The explicit `--kb` option takes precedence over the configured default.

## Preserved Tools and Automated Tests

The original engineering tools and development workflows remain available.

`search_docs.py` continues to search the 70 curated documentation records.

`index_repository.py` remains available as the legacy repository inventory utility.

These tools do not replace complete source retrieval through `offline_docs.py`.

The project also preserves:

- VBA source inspection.
- Workbook inspection.
- Skill validation.
- Explicitly invoked COM testing.
- Optional verification of public documentation URLs.
- Existing development examples and helper utilities.

Public URL verification remains disabled unless `--allow-network` is explicitly specified. Network access is not required for normal offline documentation workflows.

The original eight tool/helper files remain unchanged.

To run validation and tests from the skill directory:

```powershell
py -3 scripts\validate_skill.py

py -3 -m unittest discover -s tests -v

py -3 tests\verify_snapshot.py --archive "<original-VBA-Docs-main.zip>" --out "<new-report.json>"
```

The final command is optional. It compares every file in the knowledge base against the original input archive.

The automated tests were executed in a Linux environment.

Some filesystem-link tests may require additional privileges on Windows.

The test verifier's `--enforce-os` option requires Linux and `libseccomp`. The documentation search and read tools do not require this library.

## Scope and Limitations

**Complete documentation coverage** means all documents included in the supplied Microsoft VBA-Docs repository snapshot.

It does not mean:

- Coverage of every possible VBA API.
- Independently verified synchronization with the latest Microsoft GitHub repository.
- Guaranteed correctness or completeness of the original Microsoft documentation.
- Guaranteed compatibility with every AI coding agent.
- Successful compilation or execution of generated VBA code.
- Completed end-to-end testing of Microsoft Excel, COM automation, or UserForms.

Any omissions, inconsistencies, or historical inaccuracies in the original Microsoft documentation remain visible rather than being silently corrected.

External Win32, Visual Studio Code, and pywin32 documentation is not included in the Microsoft VBA-Docs repository snapshot.

### Offline Operation vs. Offline AI Models

The skill does **not** disable network access for the entire AI client.

It also does not convert a cloud-based AI model into a locally running model.

The documentation search, retrieval, and verification utilities can operate without internet access.

Whether the AI model itself can operate fully offline depends on the selected model, client, and configuration.

### Production Safety

AI-generated VBA code must be reviewed and tested before deployment to production workbooks or business-critical environments.

The skill encourages controlled changes, source-backed API usage, explicit error handling, and regression testing, but it cannot eliminate all implementation risks.

## Documentation Sources and Licensing

The offline documentation corpus is derived from Microsoft's publicly available [MicrosoftDocs/VBA-Docs repository](https://github.com/MicrosoftDocs/VBA-Docs).

The original Microsoft documentation, applicable licenses, copyright notices, and third-party notices are preserved in the knowledge package.

Microsoft documentation and project-specific engineering instructions are maintained as separate components.

This project is independent and is **not affiliated with, endorsed by, or sponsored by Microsoft**.

For detailed attribution and licensing information, see [LICENSES.md](LICENSES.md) and [SOURCES.md](SOURCES.md).

## Project Documentation

| Document | Description |
|---|---|
| [SOURCES.md](SOURCES.md) | Documentation sources and provenance |
| [QUALITY-REPORT.md](QUALITY-REPORT.md) | Documentation quality and indexing information |
| [ACCEPTANCE.md](ACCEPTANCE.md) | Acceptance criteria and validation evidence |
| [LIMITATIONS.md](LIMITATIONS.md) | Known technical limitations |
| [LICENSES.md](LICENSES.md) | Licensing, copyright, and attribution |
| [CHANGELOG.md](CHANGELOG.md) | Project version history |
| [OFFLINE-VERIFICATION.md](OFFLINE-VERIFICATION.md) | Recorded offline documentation retrieval tests |

## Feedback and Contributions

Excel VBA Expert is publicly available for developers and AI coding agent users interested in reliable, source-backed VBA engineering.

Feedback, bug reports, compatibility testing, and contributions are welcome through GitHub Issues and Pull Requests.

**Project repository:** https://github.com/cycu767/excel-vba-expert

**Latest release:** https://github.com/cycu767/excel-vba-expert/releases

**Official Microsoft Excel VBA documentation:** https://learn.microsoft.com/en-us/office/vba/api/overview/excel
