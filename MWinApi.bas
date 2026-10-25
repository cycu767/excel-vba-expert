Attribute VB_Name = "MWinApi"
Option Explicit
' Declarations only. Do not call Sleep for UI responsiveness.
' HWND is pointer-sized; Sleep DWORD milliseconds remains Long on Win64.
#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Function ActiveWindowHandleText() As String
    ' This is the calling thread's active window, NOT a promise of the Excel main HWND.
    ActiveWindowHandleText = CStr(GetActiveWindow())
End Function
