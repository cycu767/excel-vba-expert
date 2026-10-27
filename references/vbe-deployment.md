# VBE, references and deployment boundaries

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Establish the real host
[ENG] Record installed Excel build/channel, VBA/Office bitness, references and required ActiveX components. Do not infer a missing library only from a generic compile error. Inspect Tools > References in the actual project and resolve the intended dependency; do not indiscriminately remove every MISSING item.

[DOC: MS-LANGUAGE, MS-DEBUG-MENU] The language documentation includes the VBE add-in model and UI help, and VBE provides project compilation. [DOC] The local `Language/Reference/Visual-Basic-Add-in-Model/` corpus includes
five grouped chapters (objects, collections, properties, methods and events).
For example, `VBProject` and `VBComponent` appear under headings in
`objects-visual-basic-add-in-model.md`. Do not mistake the absence of a
standalone `api/VBIDE.VBComponent.md` for missing documentation.
[ENG] These grouped pages are not an exhaustive, per-object VBIDE member
inventory. Run `api_preflight.py --namespace vbide --symbol VBComponent`, read
its original lines, and check membership and version in the target Object
Browser before implementing editor automation.

## Preserve component identity
[ENG] Export standard modules, classes and forms using an approved workflow. Keep hidden attributes, document-module identity and .frm/.frx pairing. An imported class called Sheet1 is not automatically the host's original sheet document module. Do not delete/recreate worksheets as a shortcut that loses event binding or CodeName identity.

Check encoding of exported code and preserve localized strings. Compare source diffs and compile after import. A text extraction from vbaProject.bin is not a lossless export of every designer resource or project setting. Never use a macro container writer as an excuse to bypass project protection or trust settings.

## Declared code page, not a guessed one
[ENG] Legacy VBA storage declares its code page inside the project's own metadata,
not in the module text. When reading module streams directly, take the code page
from that declaration. Do not fall back to a plausible default such as cp1252 when
the project declares a different one, and do not treat "it decoded" as proof the
decoding was right.

[ENG] A wrong single-byte code page rarely raises. It produces text that still
looks like valid source while silently corrupting every non-ASCII character, so
identifiers, string literals and comments change meaning without any error. If a
decoded project contains unexpected characters, verify the declared code page
before concluding the source is damaged or that a re-encode is needed.

[ENG] When comparing two projects, decode both with their own declared code page
and report the code page used. Comparing a cp1250 project decoded as cp1252
against a correct read produces differences that are pure decoding artefacts, and
comparing two identically wrong reads produces a false match.

[ENG] Import is a separate encoding hazard from reading. An importer that decodes a
source file using the system code page will destroy a file saved as UTF-8: every
multi-byte character becomes mojibake, and part of it becomes U+FFFD, which is
unrecoverable from the imported result. Fix it on both sides — write exported and
generated files in the code page the importer expects, and make the importer
reject a file carrying a BOM, containing U+FFFD, or otherwise looking like UTF-8
masquerading as ANSI. Keep the pre-import copy so damaged positions can be restored
by aligning against undamaged text rather than retyping it.

[ENG] Detect this by counting replacement characters before and after the import,
not by reading the result and hoping to notice. Zero before and non-zero after is
proof the import, not the source, caused the damage.

## Verifying VBA without Excel and without Trust access
[ENG] `Workbook.VBProject` requires "Trust access to the VBA project object
model". That setting is a security boundary: do not enable it, and do not accept
a suggestion to enable it, merely to read code for comparison. When the only goal
is to prove that project code is unchanged, read the source out of the package
instead. `compare_vba_source.py` decompresses each module stream from
`vbaProject.bin` with no Excel, no Trust-Center change and no third-party
package.

[ENG] Identical `vbaProject.bin` bytes prove byte-for-byte identity of that part,
including its contained code; a **difference** in those bytes does not prove that
code changed. Excel may re-serialise the OLE binary on save. To investigate a
binary difference, compare extracted component sources separately and report
source equality without upgrading it to "the entire project is unchanged".

[ENG] Source equality has a hard ceiling. It does not cover designer resources
behind a form's control layout, project properties, reference bindings, hidden
attributes not present in the module text, or compiled behaviour. Where those
matter, say which of them remain unverified and name the test that would cover
them.

[ENG] A decoder written for this purpose must be cross-validated against an
independent reader before its verdict is used, and the cross-validation must
assert non-vacuity: component count and total source lines above a plausible
floor, and an equal number of matched components. Comparing two empty results
looks like a pass.

## Security and deployment
[ENG] Keep macro policy, trusted locations and signing under the organization's control. Do not automatically enable programmatic access to VBProject. Use an authorized signing/deployment process after code changes and verify trust on a clean target machine. Test optional library availability and provide a clear diagnostic/fallback only when its semantics are equivalent.

The installed skill is passive guidance. It cannot make a client support skills or grant Windows automation capabilities. In VS Code, validate discovery using the actual agent harness; Qwen model selection alone does not establish skill loading. Follow README.md for placement and explicit invocation tests.

## Acceptance
[ENG] Clean-machine compile; reference resolution; existing button/event callbacks; form import and designer resources; 32/64-bit support; signing/trust behavior; installation discovery; rollback of the release copy. Report each unavailable test as NOT_RUN, not as a successful deployment.

## Source register

- [MS-LANGUAGE: VBA language reference](https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference)
- [MS-DEBUG-MENU: Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu)
- [MS-64BIT: 64-bit VBA overview](https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview)
- [VSC-SKILLS: VS Code Agent Skills](https://code.visualstudio.com/docs/agent-customization/agent-skills)
