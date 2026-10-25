# Diagnose a VBA defect from evidence

All steps below are engineering workflow rules, not claims that tests have already run.
## Capture the failure
Record exact error number/text, failing line, procedure/module, inputs, workbook
names, Excel build/bitness and reproduction steps. Ask only for missing evidence
that cannot be read from the project. Do not request confidential production data
when a reduced fixture reproduces the issue. Preserve a failing copy.

## Branch on evidence
- Compile error: run Debug > Compile when available; inspect declarations,
  conditional compilation, missing references and duplicate procedure names first.
- Runtime error: stop near the first failure rather than at a downstream Nothing.
  Capture Err.Number/Description/Source before another operation overwrites them.
- Wrong result without error: define the expected result from a trusted fixture;
  compare input, intermediate values, types, bounds and final output.
- Event loop or hang: inspect re-entry guards, recursive writes, DoEvents, modal
  dialogs and background refresh. Do not add more DoEvents as a default remedy.
- COM failure: use the COM workflow; distinguish external client errors from the
  VBA handler and identify which process/workbook the client owns.

## Instrument, hypothesize, isolate
Use a breakpoint before the suspect operation, Step Into/Over/Out, Locals and
Call Stack. Use Immediate/Watch only for expressions without unintended side
 effects: property getters and function calls can execute code. For external COM
runs, avoid break-mode interaction that makes the client wait on an invisible IDE.
Add bounded, redacted Debug.Print tracing at entry/exit and before critical writes.
Use Debug.Assert for interactive invariants, not a batch test completion signal.

Maintain a short hypothesis table: proposed cause, discriminating observation,
actual observation, accepted/rejected/unknown. Reduce the failing input, not the
business contract. Check qualified object names and Variant subtypes before
changing algorithms. For Worksheet_Change test multi-cell paste and recursion;
for UserForm events test handler lifetime and repeated activation.

## Minimal correction and regression
Explain the causal chain with file/procedure/line evidence. Patch only the confirmed
cause or label an experiment explicitly. Re-run the original reproduction, adjacent
boundary cases and a previously passing path. Check final application state and
workbook identity. If runtime access is absent, deliver the exact probe/test for the
user; label the conclusion provisional, not a verified fix.
