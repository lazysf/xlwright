Attribute VB_Name = "S14Module"
Option Explicit

Public Sub SetInput21()
    ThisWorkbook.Worksheets("Probe").Range("B2").Value2 = 21
End Sub

Public Sub SetManual31()
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Worksheets("Probe").Range("B2").Value2 = 31
End Sub

Public Sub SetAutomatic41()
    ThisWorkbook.Worksheets("Probe").Range("B2").Value2 = 41
    Application.Calculation = xlCalculationAutomatic
End Sub
