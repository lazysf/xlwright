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

Public Sub AskOkCancel(control As IRibbonControl)
    MsgBox "続行しますか", vbOKCancel
End Sub

Public Sub Quiet(control As IRibbonControl)
End Sub

Public Sub DoneTwice(control As IRibbonControl)
    MsgBox "1 つ目の通知"
    MsgBox "2 つ目の通知"
End Sub

' 10 秒動いてから終わる。ブックを閉じる時に待つかを見る
Public Sub RunLong(control As IRibbonControl)
    Application.Wait Now + TimeSerial(0, 0, 10)
End Sub

Public Sub GetOff(control As IRibbonControl, ByRef returnedVal)
    returnedVal = False
End Sub

Public Sub GetText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Dyn"
End Sub

Public Sub Spin()
    Do
        DoEvents
    Loop
End Sub
