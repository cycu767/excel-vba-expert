# Formulas, numeric precision and localization

Evidence labels: **[DOC]** documented behavior; **[ENG]** skill engineering rule; **[INFERENCE]** derived interpretation; **[UNVERIFIED]** needs target-environment verification.

## Formula representation
[DOC: MS-FORMULA2] Formula2 uses array evaluation and may spill; Formula remains for compatibility. [ENG] Choose Formula, Formula2, FormulaR1C1 or localized variants according to target capabilities and API contract, not by string replacement of translated function names. Verify individual variants in official documentation/Object Browser before using them; this core index does not cover every formula property.

Record whether a transformation must preserve formulas, values or display text. Value2 transfer is not formula-preserving copy. Identify spilled/array formulas before overwriting part of their output. Test supported Office builds rather than assuming a version family implies every dynamic-array feature.

## Input parsing and display
[DOC: MS-VALUE2, MS-TYPES] Value2 does not use Date/Currency subtypes; VBA numeric types have distinct precision/range. [ENG] Separate user text, normalized domain values, workbook values and formatted display. Validate decimal/group separators, signs and date ambiguity explicitly. Do not parse a monetary input by replacing all commas with periods or stripping every non-digit character.

Use explicit date construction for unambiguous values; test workbook date-system assumptions when comparing serials across files. Preserve account/product identifiers as text. Choose a policy for formulas supplied in imported text and do not inadvertently turn untrusted strings beginning with formula characters into executable workbook formulas.

## Calculation contract
[ENG] Identify which results must be current before save/export. Manual mode can leave cached results stale. Recalculate the necessary scope under an explicit policy and validate the outputs; do not assume opening a file or calling a Python spreadsheet library recalculates Excel formulas.

For money, specify scale, rounding rule, rounding stage and accumulation order with the business owner. Do not silently change existing rounding during a performance refactor. For floating point, choose a comparison based on the domain magnitude and expected error; for exact identifiers and counts, require exact agreement.

## Tests
[ENG] Polish and English locale; comma decimal separator; date strings that are ambiguous in another locale; leap dates; negative/zero/large values; leading zeros; formula errors; blank versus zero; spill collisions; stale manual-calculation results; changed formula semantics under Formula2.

## Source register

- [MS-FORMULA2: Range.Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2)
- [MS-VALUE2: Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
- [MS-TYPES: Data type summary](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary)
- [MS-PERFORMANCE: Excel performance obstructions](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)
