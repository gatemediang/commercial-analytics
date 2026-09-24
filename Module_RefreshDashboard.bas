Attribute VB_Name = "Module_RefreshDashboard"
Option Explicit

' Refreshes workbook calculations, connections and PivotTables where available.

Public Sub RefreshCommercialModel()
    Dim ws As Worksheet
    Dim pt As PivotTable

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic

    ThisWorkbook.RefreshAll
    Application.CalculateFullRebuild

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.RefreshTable
            On Error GoTo 0
        Next pt
    Next ws

    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Range("B14").Value = _
        "Last refreshed: " & Format(Now, "dd-mmm-yyyy hh:mm")
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Commercial model refresh complete.", vbInformation
End Sub
