# Windows API and VBA7/64-bit declarations

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Review the native signature first
[DOC: MS-64BIT] PtrSafe is required for 64-bit declarations but does not correct the data types. LongPtr is for pointers/handles; Long stays 32-bit. [ENG] For each parameter and return value, build a table: native type, width/meaning, ByVal/ByRef, VBA type, lifetime/ownership and API source. Check every variable and UDT field that stores the value, not only Declare.

Distinguish VBA7 (language features) from Win64 (Office host bitness). Neither the Windows installation nor Python's pointer size alone determines Office bitness. Preserve a legacy branch only if the supported host matrix actually includes pre-VBA7 versions. Do not mechanically replace every Long with LongPtr: DWORD and BOOL are not pointer-sized handles.

## Verified contrast
[DOC: MS-GETACTIVEWINDOW, MS-SLEEP] GetActiveWindow returns HWND; Sleep accepts a DWORD count and returns no value. [ENG] The VBA7 declarations in `MWinApi.bas` use LongPtr for HWND and Long for a nonnegative practical Sleep duration. The template includes a VBA6 branch for illustration. Do not call Sleep from the UI to fake asynchronous work, or confuse GetActiveWindow with the foreground window of another process.

## Strings, structures and callbacks
[ENG] Confirm ANSI/Unicode entry point and string marshaling; do not infer a W-suffixed declaration by appending W to a copied ANSI declaration. Review buffers, lengths and ownership. For UDTs, validate native layout/alignment and Len/LenB expectations against the real API. Audit callbacks, AddressOf, subclassing and stored pointers for lifetime and host shutdown behavior. These advanced patterns are NOT shipped as tested templates here.

Prefer Excel's supported object model when it meets the requirement. Window-handle tricks for modeless forms, DPI changes or message pumping add failure modes and need focused Windows tests. Do not patch process memory or install global hooks for cosmetic UI work.

## Acceptance matrix
[ENG] Compile supported 32-bit and 64-bit Office configurations; verify returned handles are not truncated, zero/error cases are handled, and the surrounding code stores the correct width. Test the feature and close/unload paths. A syntactic PtrSafe scan is only a candidate detector. If the native signature or target bitness is unknown, stop before changing the declaration and mark it unresolved.

## Source register

- [MS-64BIT: 64-bit VBA overview](https://learn.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview)
- [MS-GETACTIVEWINDOW: GetActiveWindow (Win32)](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getactivewindow)
- [MS-SLEEP: Sleep (Win32)](https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-sleep)
