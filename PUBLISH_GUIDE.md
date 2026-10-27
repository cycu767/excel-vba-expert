# GitHub publication checklist - Excel VBA Expert v2.1.0

Repository: https://github.com/cycu767/excel-vba-expert

## Files in this publication kit

- `README.md`: complete, English, professional GitHub landing page for v2.1.0.
- `CHANGELOG.md`: v2.1.0 plus release history through v1.0.0.
- `RELEASE_NOTES.md`: text to paste into the new GitHub Release description.
- `excel-vba-expert-v2.1.0.zip`: publishable core skill with one `excel-vba-expert/` top-level folder.
- `SHA256SUMS-v2.1.0.txt`: SHA-256 of that core ZIP.

## 1. Update repository files first

Start from an up-to-date clone of the `main` branch. Extract the *content* of
`excel-vba-expert-v2.1.0.zip` and copy its `excel-vba-expert/` contents to the
repository root; do NOT create a nested `excel-vba-expert/` directory in Git.
Overwrite matching source files and documentation. The release ZIP already
contains the new `README.md`, `CHANGELOG.md` and refreshed integrity metadata.

**Preserve the existing GitHub `assets/readme-hero.svg` from v2.0.2.** The core
ZIP deliberately does not include the old banner artwork. The new README
links to the existing hosted version so the project keeps the same visual look.
Do not delete the banner as an 'extra' file.

If the old repository still contains `tests/last-run.txt`, remove it: it is an
outdated test log from a previous release and not part of the current package.
Do NOT upload the retrospective archive, internal analysis, patches, project
files, Microsoft documentation database, or personal paths.

Before commit, run at the repository root:

```powershell
python scripts\validate_skill.py --root .
```

Expected output: PASS, no warnings, 20 modules, 20 scenarios, 14 scripts.
The banner is an additional existing asset in GitHub and is retained.

Commit suggestion:

`release: Excel VBA Expert v2.1.0 - evidence-driven VBA engineering`

Push the commit to `main`, then inspect the rendered README: badges, banner,
Mermaid charts, new feature links, and the v2.1.0 download link (which will
work as soon as the release is published).

## 2. Publish the GitHub Release

Go to Repository > Releases > Draft a new release.

- Tag: `v2.1.0` (create from the updated `main` commit).
- Release title: `Excel VBA Expert v2.1.0 - Evidence-Driven VBA Engineering`.
- Description: paste the complete contents of `RELEASE_NOTES.md`.
- Release assets: attach `excel-vba-expert-v2.1.0.zip` and `SHA256SUMS-v2.1.0.txt`.
- The separate `excel-vba-knowledge.zip` was not changed; link to the v2.0.2
  release rather than reuploading an unverified copy.

Publish only after the new tag points to the correct commit and the assets are
present. Do not attach the retrospective archive.

## 3. Quick release verification

- Open the `v2.1.0` release and verify both files download.
- Verify the release ZIP has `excel-vba-expert/SKILL.md` at its expected path.
- Check its SHA-256 against the published checksum file.
- Check the README badge points to v2.1.0 and its banner renders.
- Keep the existing local `excel-vba-expert-knowledge` folder unchanged.

## Evidence boundaries

Only documentation, public metadata and packaging were changed after the
maintainer's 156-test acceptance of the code. The 31 comparison-tool tests
were also executed in the publication workspace and passed. The structural
validator passed for this rebuilt ZIP. The complete KB-dependent test suite
was not rerun here because the companion knowledge package was not available
in the packaging environment; the earlier maintainer report remains the
record of 156 passed tests (1 skipped).

The tools do not compile VBA or perform live Excel execution.
