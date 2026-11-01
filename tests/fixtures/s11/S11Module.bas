Attribute VB_Name = "S11Module"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Sub S11FiniteWait()
    Sleep 15000
End Sub

Public Sub S11Modal()
    MsgBox "xlwright S11 modal", vbOKOnly, "xlwright S11"
End Sub
