Attribute VB_Name = "S10Module"
Option Explicit

Public Sub Auto_Open()
    S10Trace "auto-open"
End Sub

Public Function S10Probe(ByVal x As Double) As Double
    S10Probe = x + 1
End Function

Public Sub S10Trace(ByVal tag As String)
    Dim channel As Integer
    channel = FreeFile
    Open ThisWorkbook.Path & "\s10-vba.log" For Append As #channel
    Print #channel, tag
    Close #channel
End Sub
