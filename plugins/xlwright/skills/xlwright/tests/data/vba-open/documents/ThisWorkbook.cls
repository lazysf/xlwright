VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 開かれた印に、ブックと同じフォルダに opened.txt を書く
Private Sub Workbook_Open()
    Dim handle As Integer
    handle = FreeFile
    Open Me.Path & "\opened.txt" For Output As #handle
    Print #handle, "opened"
    Close #handle
End Sub
