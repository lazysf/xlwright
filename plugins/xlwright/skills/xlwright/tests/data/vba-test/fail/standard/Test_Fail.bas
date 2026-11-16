Attribute VB_Name = "Test_Fail"
Option Explicit

Public Sub Test_1_Pass()
End Sub

Public Sub Test_2_Raise()
    Err.Raise vbObjectError + 1, , "expected 3, actual 2"
End Sub

Public Sub Test_3_PassAfterFailure()
End Sub

Public Sub Test_4_Note()
    MsgBox "note", vbOKOnly, "probe"
End Sub

Public Sub Test_5_Ask()
    MsgBox "question", vbYesNo, "probe"
End Sub

Public Sub Test_6_Never()
End Sub
