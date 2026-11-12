Attribute VB_Name = "Util"
Option Explicit

' 2 つの整数を足す
Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function

Public Function Greeting() As String
    Greeting = "こんにちは、ｶﾅ"
End Function

Public Sub Done(control As IRibbonControl)
    MsgBox "処理が完了しました"
End Sub

Public Sub AskYesNo(control As IRibbonControl)
    MsgBox "続けますか", vbYesNo
End Sub

Public Sub GetOff(control As IRibbonControl, ByRef returnedVal)
    returnedVal = False
End Sub

Public Sub Spin()
    Do
        DoEvents
    Loop
End Sub
