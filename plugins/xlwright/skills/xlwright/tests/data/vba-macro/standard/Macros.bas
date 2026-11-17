Attribute VB_Name = "Macros"
Option Explicit

Private calls As Long

Public Function AddThree(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    AddThree = a + b + c
End Function

Public Sub CountCall()
    calls = calls + 1
End Sub

Public Function CallCount() As Long
    CallCount = calls
End Function

Public Function Slow() As Long
    Application.Wait Now + TimeSerial(0, 0, 2)
    Slow = 7
End Function

Public Function DivZero() As Long
    Dim zero As Long
    DivZero = 1 \ zero
End Function

Public Sub DivZeroSub()
    Dim zero As Long
    calls = 1 \ zero
End Sub

Public Function ShowTwice() As Long
    MsgBox "first message", vbOKOnly, "probe"
    MsgBox "second message", vbOKOnly, "probe"
    ShowTwice = 5
End Function

Public Function Ask() As Long
    Ask = MsgBox("question", vbYesNo, "probe")
End Function

Public Function AskOkCancel() As Long
    AskOkCancel = MsgBox("proceed", vbOKCancel, "probe")
End Function
