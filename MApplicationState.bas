Attribute VB_Name = "MApplicationState"
Option Explicit
' Original engineering template. NOT compiled in the delivered build.
' Copies values only. Does NOT roll back data or save any workbook.
Private Type TApplicationState
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    Valid As Boolean
End Type

Public Sub CopyValuesSafely(ByVal source As Excel.Range, _
                            ByVal target As Excel.Range, _
                            Optional ByVal useManualCalculation As Boolean = False)
    Dim app As Excel.Application
    Dim prior As TApplicationState
    Dim originalNumber As Long
    Dim originalSource As String
    Dim originalDescription As String
    Dim cleanupMessage As String

    On Error GoTo Failed
    If source Is Nothing Then Err.Raise 5, "CopyValuesSafely", "Source is required."
    If target Is Nothing Then Err.Raise 5, "CopyValuesSafely", "Target is required."
    If Not (source.Application Is target.Application) Then _
        Err.Raise 5, "CopyValuesSafely", "Use ranges from the same Excel instance."
    If source.Areas.Count <> 1 Or target.Areas.Count <> 1 Then _
        Err.Raise 5, "CopyValuesSafely", "One rectangular area is required."
    If source.Rows.Count <> target.Rows.Count Or _
       source.Columns.Count <> target.Columns.Count Then _
        Err.Raise 5, "CopyValuesSafely", "Source and target dimensions differ."

    Set app = source.Application
    prior.EnableEvents = app.EnableEvents
    prior.ScreenUpdating = app.ScreenUpdating
    prior.DisplayAlerts = app.DisplayAlerts
    prior.Calculation = app.Calculation
    prior.Valid = True  ' No setting has changed until the complete snapshot exists.

    app.EnableEvents = False
    app.ScreenUpdating = False
    If useManualCalculation Then app.Calculation = xlCalculationManual
    ' DisplayAlerts deliberately stays unchanged; suppression is not a safety feature.
    target.Value2 = source.Value2
    GoTo CleanUp

Failed:
    originalNumber = Err.Number
    originalSource = Err.Source
    originalDescription = Err.Description
    Resume CleanUp  ' Leave active handler state before any propagation.

CleanUp:
    On Error GoTo 0
    If prior.Valid Then cleanupMessage = RestoreApplication(app, prior)
    Set app = Nothing
    If originalNumber <> 0 Then
        If Len(cleanupMessage) > 0 Then originalDescription = originalDescription & _
            " | Cleanup failures: " & cleanupMessage
        Err.Raise originalNumber, originalSource, originalDescription
    End If
    If Len(cleanupMessage) > 0 Then _
        Err.Raise vbObjectError + 2048, "CopyValuesSafely.Cleanup", cleanupMessage
End Sub

Private Function RestoreApplication(ByVal app As Excel.Application, _
                                    ByRef prior As TApplicationState) As String
    Dim details As String
    ' Scoped suppression: every restoration attempted, every error recorded.
    On Error Resume Next
    Err.Clear
    app.Calculation = prior.Calculation
    AddCleanup details, "Calculation", Err.Number, Err.Description
    Err.Clear
    app.DisplayAlerts = prior.DisplayAlerts
    AddCleanup details, "DisplayAlerts", Err.Number, Err.Description
    Err.Clear
    app.ScreenUpdating = prior.ScreenUpdating
    AddCleanup details, "ScreenUpdating", Err.Number, Err.Description
    Err.Clear
    app.EnableEvents = prior.EnableEvents  ' Last; calculation restored with events off.
    AddCleanup details, "EnableEvents", Err.Number, Err.Description
    Err.Clear
    On Error GoTo 0
    RestoreApplication = details
End Function

Private Sub AddCleanup(ByRef details As String, ByVal setting As String, _
                       ByVal number As Long, ByVal description As String)
    If number = 0 Then Exit Sub
    If Len(details) > 0 Then details = details & "; "
    details = details & setting & ": " & CStr(number) & " " & description
End Sub
