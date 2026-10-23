Attribute VB_Name = "MValidation"
Option Explicit
' Small pure function for unit testing independent of worksheets and UserForms.
' Contract: ASCII unsigned integer only, 1..1000000. No locale-dependent parsing.
Public Function TryParseQuantity(ByVal text As String, ByRef quantity As Long) As Boolean
    Dim trimmed As String
    Dim i As Long
    Dim digit As Long
    Dim result As Long
    quantity = 0
    trimmed = Trim$(text)
    If Len(trimmed) = 0 Or Len(trimmed) > 7 Then Exit Function
    For i = 1 To Len(trimmed)
        digit = AscW(Mid$(trimmed, i, 1)) - AscW("0")
        If digit < 0 Or digit > 9 Then Exit Function
        result = result * 10 + digit
        If result > 1000000 Then Exit Function
    Next i
    If result < 1 Then Exit Function
    quantity = result
    TryParseQuantity = True
End Function
