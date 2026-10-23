# Offline knowledge acceptance matrix - version 2.0.0

| Requested gate | Status | Evidence |
|---|---|---|
| Extend existing excel-vba-expert, not recreate it | PASS | Preservation manifest: 19 modules, 7 workflows, 8 old Python files unchanged |
| Import real Microsoft VBA-Docs sources | PASS | Supplied ZIP SHA-256 and per-file comparison |
| Complete local directory/ZIP import process | PASS | import_knowledge.py plus ZIP/directory/security tests |
| Physical full documentation and licenses | PASS | 38,911 files, including 38,187 Markdown and original license notices |
| Index local full text, not URL summaries | PASS | SQLite FTS5 plus local document paths in JSONL |
| Objects/methods/properties/events/enums lookup | PASS | Exact aliases, category/kind filters, constants and full-text tests |
| Five named API documents readable | PASS | Full source/hash/section evidence in offline-verification.json |
| Mandatory local-first workflow | IMPLEMENTED | SKILL.md; actual live-model compliance not measured |
| Expert/source layers cooperate | PASS | 62 curated IDs crosswalk to full local Markdown; original modules retained |
| Automatic tests with networking disabled | PASS - 117/117 | Kernel seccomp runner; actual counts in offline-test-results.json |
| Actual imported/indexed/local counts | PASS | All three Markdown counts are 38,187 |
| Package size and portable paths | TWO-PACKAGE DELIVERY | Small skill.zip plus separate knowledge ZIP; sibling paths |
| Attribution and original license files retained | PASS | LICENSE, LICENSE-CODE, ThirdPartyNotices, source attributions and separate authored indexes |
| Search/read from moved installation and unrelated CWD | TESTED | Relative config and relocation tests; final archive check recorded separately |
| Ask a real offline Qwen/Copilot agent and evaluate generated VBA | NOT_RUN | Requires target agent/client; not claimed from tool tests |
| Compile/run generated VBA in Excel | NOT_RUN | No real Excel runtime in this environment |

The retrieval/source-storage portion of the offline criterion is demonstrated.
The installed AI agent's adherence and code behavior require a local acceptance
run; neither is falsely certified here. Use the README's exact offline prompt and
record retrieved source paths, applied rules, proposed code and runtime status.
