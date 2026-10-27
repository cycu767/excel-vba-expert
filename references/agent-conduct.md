# Agent and subagent conduct on long VBA tasks

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

This module governs how an agent drives itself and delegated workers across a
multi-stage task on a real workbook. It is not about VBA code quality.

## Separate the reference point from the working copy
[ENG] Before any mutation, fix one immutable reference point (the untouched
original, read-only) and one working copy. Every later claim of "unchanged" is
made against the reference point, never against a memory of what the file used to
look like. If the reference point can be overwritten, it is not a reference point.

[ENG] A legitimate configuration change makes the working copy differ from an
older approved baseline. That difference is expected, not a defect. Do not treat a
fingerprint mismatch against a superseded baseline as a failure, and do not
"fix" it by rewriting the baseline. Record which artefact is the reference point
and which is superseded, in writing, before the change.

## Verification must run after the work it verifies
[ENG] A check that aborts on a precondition inherited from an earlier phase has
verified nothing. If a gate fires before the substantive comparisons run, the exit
code describes the gate, not the work. Re-derive the gate for the current
situation, or split the check so the parts that can run do run, and report exactly
which parts ran.

[ENG] Never report a blocked or aborted verification as a pass, and never let a
tool's silence stand in for a result. If a check cannot be executed, record it as
NOT_RUN or BLOCKED with the specific blocker and the next executable step.

## Delegation on a restricted harness
[ENG] Verify a delegated worker actually runs before planning around it. A worker
that fails to start, fails on model availability, or returns nothing is not a
result. Confirm the first successful round trip, then delegate; do not assume a
configured model or endpoint is reachable.

[ENG] Give each worker a bounded read-only objective with an explicit stop
condition, and consume its report rather than repeating its work. Do not relaunch a
worker that returned nothing useful; continue directly with the evidence already
in hand. Do not let a worker's summary replace verification that only a test or a
comparison can provide.

[ENG] Long tasks stall visibly: a worker can run for many minutes without output.
Distinguish "working", "waiting" and "lost". While waiting, do independent work
that does not depend on the answer. Do not poll in a tight loop, and do not
interpret delay as failure.

## Do not relax security to unblock a step
[ENG] When a read is blocked by a security setting, find a read path that does not
require the setting. Reading VBA source offline instead of enabling programmatic
project access is the pattern: the obstacle is a requirement to use a different
tool, not a reason to lower a protection. This includes Trust access to the VBA
project model, macro settings, protection removal and signing changes.

[ENG] Refusing to change a security setting is a deliverable, not a failure to
complete. State the blocked route, the substitute route taken, and what the
substitute can and cannot prove.

## Stop where the specification says stop
[ENG] When the task defines a stop condition ("halt and present the design if a new
guard unlock is needed"), reach it, measure the blocking items rather than assume
them, and do not implement until the owner decides. Where several decisions are
required, state the order they must be taken in and respect that order.

[ENG] Close every stage at the boundary of the approved scope and write down what
was *not* done and why. An approved decision that was deliberately not implemented
must be labelled unimplemented with the list of missing artefacts. Silently dropping
it, or implying it shipped, is worse than reporting it unfinished.

[ENG] An acceptance decision states its own scope and evidential basis. Acceptance
of one phase is not acceptance of the whole stage, and neither is an independent
audit of the source. Say which.

[ENG] Do not re-run a closed verification merely to confirm a decision when the
code, relevant inputs, configuration and environment have not changed since the
tests. Reuse the existing evidence and record why it still applies. A change in
any of those conditions, or a newly discovered defect, can justify targeted or
full re-verification according to the risk; business-choice approval alone does not.

## A long build must be resumable
[ENG] A multi-step build that can be interrupted must be able to resume: skip the
steps already completed, but re-verify the current state before executing the
remaining ones. A fail-closed guard that refuses to restart because an artefact
already exists turns an interruption into a dead end; a resume mode is part of the
design, not a later convenience.

[ENG] Audit stages produce evidence and a plan, not changes. Dead code confirmed
unreachable stays until a separate approved cleanup runs, in batches, with a
"do not touch" category. Verification that removal is safe (compiles, output
bit-identical) is not authorization to remove.

[ENG] Before moving or deleting a file, scan for references to its exact name and
classify the dependencies you find. After cleanup, run a dependency check over the
whole tool set. Files that an installer points at directly are not orphans even if
nothing else references them.

## When your own edit breaks something
[ENG] If an edit you made caused the damage, say so and quantify it (for example
"before my edit the module had 0 replacement characters, after it had 21"). Keep a
pre-edit copy from which the damaged positions can be restored, and fix the cause
systemically: make the importer reject a file with a BOM, a replacement character,
or UTF-8 masquerading as the declared single-byte code page.

[ENG] Never run a generator, formatter or installer without first reading what it
writes. A generator that embeds its own stale copy of the source will overwrite the
current file. If the source survives only inside a workbook, export it before
running anything that could clobber it.

[ENG] If a conclusion rested on a stale artefact, re-derive it from the actually
installed state and withdraw the earlier claim explicitly. If a decision was given
wrongly and then corrected, verify that the wrong one had not already taken effect
before acting on the correction.

[ENG] When a gate starts failing after a deliberate change, first ask whether the
gate still points at the right artefact. A comparison of pre-change code in an old
carrier against post-change sources must fail; that is a stale test target, not an
application regression. Re-point the gate, and do not weaken the assertion to make
it green.

## Report scope, not conclusions
[ENG] For every claim, state the layer it came from and the scope it covers.
"VBA source identical across 66 components" is a claim; "the workbook is
unchanged" is a different, broader claim that the first does not support. Name the
sheets, components and properties compared and those not compared.

[ENG] When asked to analyse without changing anything, produce analysis only.
Resist the pull to "just fix" a defect discovered on the way; record it as a
finding with its evidence and let the owner decide. An unrequested change inside a
read-only task is a scope violation even when the change is an improvement.

## Source register

- [VBProject object](https://learn.microsoft.com/en-us/office/vba/language/Reference/Visual-Basic-Add-in-Model/vbproject-object)
- [MS-AUTOMATION-SECURITY: Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
- [VSC-SKILLS: VS Code Agent Skills](https://code.visualstudio.com/docs/agent-customization/agent-skills)
