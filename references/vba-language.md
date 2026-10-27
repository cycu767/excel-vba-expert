# VBA language engineering

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Types and contracts
[DOC: MS-EXPLICIT, MS-DIM, MS-TYPES] Require Option Explicit and declare each variable with its own type; `Dim a, b As Long` does not type a as Long. Long is 32-bit even under 64-bit Office; Integer is too small for general worksheet row counts. [ENG] Use Long for ordinary indexes and an appropriate wider count representation for CountLarge-scale work. Use concrete Excel/MSForms types when dependencies are present and intentional late binding at optional library boundaries.

[DOC: MS-SUB] ByRef is VBA's default. [ENG] Mark parameter intent explicitly. ByVal on an object copies its reference; it does not clone the workbook or make it immutable. Avoid ByRef mutation as a hidden secondary output. Use a Function for a result and document output parameters when actually necessary.

Differentiate Empty, Null, Nothing, empty text and CVErr values. Guard IsError and IsNull before numeric/string conversion. `IsNumeric` is not a complete business validation rule: decide accepted signs, decimal separators, exponent notation and bounds. Preserve identifier leading zeros as text.

## Expressions with side effects
[DOC: MS-IIF, MS-ASSERT] IIf evaluates both branches; boolean expressions do not provide a safe short-circuit null guard in VBA. [ENG] Replace `If obj Is Nothing Or obj.Count = 0 Then` with nested checks. Do not put writes, dialogs or expensive calls into assertion/watch expressions.

## Arrays, collections and dictionaries
[ENG] Use explicit bounds and LBound/UBound. Do not change Option Base as an incidental refactor. Handle uninitialized dynamic arrays and scalar range reads. ReDim Preserve preserves only supported dimensions; choose a growth strategy before repeated append. Treat a two-dimensional data array as a data contract with column definitions, not an unlabelled bag of offsets.

[DOC: MS-COLLECTION, MS-DICTIONARY] Collection offers ordered items and keys but not Dictionary.Exists. Dictionary offers key lookup and CompareMode. [ENG] Set key comparison policy before adding keys. Test duplicate keys and distinguish a duplicate from a component-creation failure. Late binding avoids a compile reference but does not install a missing scripting component.

## Iterating arrays returned by a function
[ENG] Iterate `LBound` to `UBound` of one stored variable, never `1 To UBound(x)`
with a manual `-1` inside the loop. That pattern silently drops the last element of
a zero-based array, and the loss is invisible whenever the output is a log or a
report rather than a computed figure. Split the text into a local array once, then
loop over its own bounds.

[ENG] Do not "fix" the missing last line by adding `+1` to `UBound`. Splitting text
that ends with the delimiter yields a trailing empty element, so `+1` writes an
empty record. Handle the empty element explicitly: skip entries that are empty
after `Trim`.

[ENG] The empty VBA array is `Array()`, not `ReDim x(0 To -1)`, which raises
run-time error 9. A helper that splits text, parses input or aggregates rows must
define behaviour for empty input, because the compiler does not check it and a
single-line test case is usually what exposes it.

## Aggregating over a dependency graph
[ENG] Roll values up a dependency graph in a dependency-first order, not simply in
table order. For child-to-parent aggregation, orient edges **child -> parent** and
run Kahn's topological sort; if edges are **parent -> child**, reverse the resulting
order. An arbitrary iteration can sum a parent before its children are ready;
recursion also requires explicit depth and cycle handling. A nonempty remainder
after Kahn proves that a cycle exists, but **not every remaining node is in a
cycle**: some may merely depend on cyclic nodes. Use strongly connected components
when you need to identify the actual cycle members. Report unresolved and blocked
nodes separately and do not include them in an ordinary sum-closure PASS.

[ENG] When several rules can match one record, define precedence explicitly and make
ties visible. "Longest/most specific match wins" is workable; an unresolved tie must
route the record to a review state with a stated reason, never to a silent winner.
Sort the rule table once at load time rather than rescanning for the best match.

## Numeric business rules
[ENG] Choose Currency only when four decimal places and its range match the domain; otherwise choose a documented Decimal/Double strategy. Do not replace calculation types, rounding or operation order during cosmetic refactoring. Specify rounding boundaries and compare monetary totals using the domain's agreed precision, not an arbitrary epsilon. Treat Excel serial dates, locale parsing and display formats separately.

## Modules and object lifetime
[ENG] Keep UI code in adapters, domain transformations in plain procedures/classes, persistence in explicit I/O boundaries. Limit Public globals. Create objects explicitly, retaining event subscribers for their required lifetime. Release cyclic references deliberately; Set Nothing alone is not a universal lifecycle solution. Do not copy .NET try/finally, generics or asynchronous patterns into VBA without a real equivalent.

## Acceptance
[ENG] Compile under the supported VBA version; test missing references, overflow, Null/error inputs, duplicate keys, non-English numeric input, zero/one/many rows, and cancellation. Static review is not compilation.

## Source register

- [MS-EXPLICIT: Option Explicit statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement)
- [MS-DIM: Dim statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
- [MS-TYPES: Data type summary](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary)
- [MS-SUB: Sub statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)
- [MS-IIF: IIf function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/iif-function)
- [MS-ASSERT: Debug.Assert](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assert-method)
- [MS-COLLECTION: Collection object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object)
- [MS-DICTIONARY: Dictionary object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)
