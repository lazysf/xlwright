Attribute VB_Name = "Test_Broken"
Option Explicit

Public Sub Test_Undeclared()
    Dim a As Long
    missing = a + 1
End Sub
