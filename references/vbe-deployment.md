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
