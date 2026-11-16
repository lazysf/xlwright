Attribute VB_Name = "Test_Calc"
Option Explicit

Public Sub Test_Add()
    If 1 + 2 <> 3 Then Err.Raise vbObjectError + 1, , "expected 3"
End Sub

Sub Test_ImplicitPublic()
End Sub

' Not tests: each of these fails when it is run.
Private Sub Test_Private()
    Err.Raise vbObjectError + 1, , "a Private Sub was run"
End Sub

Public Sub Test_WithArg(Optional ByVal value As Long = 1)
    Err.Raise vbObjectError + 1, , "a Sub with an argument was run"
End Sub

Public Function Test_Func() As Long
    Err.Raise vbObjectError + 1, , "a Function was run"
End Function

Public Sub Helper()
    Err.Raise vbObjectError + 1, , "a Sub without the prefix was run"
End Sub
