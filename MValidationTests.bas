Attribute VB_Name = "MValidationTests"
Option Explicit
' Run manually in an isolated test workbook after Debug > Compile.
Public Sub RunValidationTests()
    Dim quantity As Long
    AssertTrue TryParseQuantity(" 12 ", quantity), "Valid input"
    AssertTrue quantity = 12, "Parsed quantity"
    AssertTrue Not TryParseQuantity("", quantity), "Empty rejected"
    AssertTrue quantity = 0, "Output reset after failure"
    AssertTrue Not TryParseQuantity("0", quantity), "Zero rejected"
    AssertTrue Not TryParseQuantity("1,2", quantity), "Locale decimal rejected"
    AssertTrue Not TryParseQuantity("1e2", quantity), "Exponent rejected"
    AssertTrue TryParseQuantity("1000000", quantity), "Upper bound accepted"
    AssertTrue Not TryParseQuantity("1000001", quantity), "Over bound rejected"
    AssertTrue Not TryParseQuantity("999999999999", quantity), "Huge input rejected"
    Debug.Print "RunValidationTests: PASS (10 assertions)"
End Sub
Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2100, "MValidationTests", message
End Sub
