VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Requires Microsoft Forms 2.0, supplied by a real UserForm project.
Private WithEvents mButton As MSForms.CommandButton

Public Sub Attach(ByVal button As MSForms.CommandButton)
    If button Is Nothing Then Err.Raise 5, "CButtonHandler.Attach", "Button required."
    Set mButton = button
End Sub

Public Sub Detach()
    Set mButton = Nothing
End Sub

Private Sub mButton_Click()
    ' Demonstration only; route real commands through validated business logic.
    Debug.Print "Dynamic button clicked: " & mButton.Name
End Sub
