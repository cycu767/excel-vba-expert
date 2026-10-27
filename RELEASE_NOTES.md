# Excel VBA Expert v2.1.0 - Evidence-Driven VBA Engineering

**Safer comparisons. Stronger regression evidence. Less guessing.**

v2.1.0 builds on the public v2.0.2 release. The core skill now includes **20 expert modules**, **7 structured workflows**, **14 Python scripts** and **20 agent scenario contracts**. Its separate offline Microsoft knowledge package is unchanged.

## Highlights

### Two new read-only comparison tools

- **`compare_vba_source.py`** extracts decompressed VBA component source and compares it across macro-enabled packages without opening Excel, executing macros or enabling *Trust access to the VBA project object model*.
- **`compare_workbook_values.py`** compares worksheet cached values by **sheet and cell address**, so a targeted edit can be checked without incorrectly treating relocated values as unchanged.

Both tools use the **Python 3.10+ standard library**. They reject missing, empty or truncated inputs rather than reporting false equality. A result is always limited to the data the tool actually compares.

### Better safeguards for agents and engineers

- New `agent-conduct.md` for long-running work, reference snapshots, resumable operations, risk-aware validation and clear reporting of what was actually done.
- Stronger test design: negative and mutation tests, non-vacuous checks, detected parser failures and differentiated evidence levels.
- Better VBA guidance: `LBound`/`UBound`, safe empty arrays, qualified `Application.Run`, dependency-order processing, cycle detection and rule precedence.
- Tighter practices for encoding, offline source inspection, Excel COM ownership and verification after restoration.

## Validation

- Skill structure and source-integrity manifest: **PASS**.
- Windows maintainer acceptance: **156 Python unit tests passed; 1 skipped**, including **31 comparison-tool tests**.
- Comparison probes: **16/16 workbook cases**, **9/9 VBA cases**.
- Source extraction independently cross-checked on **two related cp1250 macro-enabled samples**; extracted component text matched in the reported scope.

**Scope matters:** this release does not claim that VBA was compiled or executed in Excel, that all MS-OVBA formats were tested, or that the 20 scenario contracts have been run against a live AI agent. Cached cell values are not formula equality; source equality is not designer-resource or runtime equality.

## Release downloads

| File | Purpose |
|---|---|
| **`excel-vba-expert-v2.1.0.zip`** | Complete core skill: instructions, modules, workflows, scripts, tests and examples. |
| **`SHA256SUMS-v2.1.0.txt`** | SHA-256 checksum for the new core ZIP. |
| [**`excel-vba-knowledge.zip` (v2.0.2)](https://github.com/cycu767/excel-vba-expert/releases/tag/v2.0.2) | Reusable offline Microsoft documentation package; **not changed or republished** for v2.1.0. |

### Updating from v2.0.2

Back up your existing `excel-vba-expert` folder, then replace its contents with the new core ZIP. **Keep `excel-vba-expert-knowledge` untouched.** Fresh installs need both the core skill and the companion knowledge archive for full offline documentation lookup.

Run:

```powershell
python scripts\validate_skill.py --root .
```

See the [README](https://github.com/cycu767/excel-vba-expert#quick-start) for complete installation instructions and verification commands.

---

**Excel VBA Expert - Less guessing. More evidence. Better VBA.**
