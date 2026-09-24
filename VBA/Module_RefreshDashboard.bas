Attribute VB_Name = "Module_RefreshDashboard"
Option Explicit

' ---------------------------------------------------------------------------
' RefreshCommercialModel
' One-click month-end refresh: queries, PivotTables, full recalculation, then
' reads the reconciliation controls on Data_Quality and stamps the result on
' the Dashboard so readers can see when the numbers were last checked.
' ---------------------------------------------------------------------------
Public Sub RefreshCommercialModel()
    Dim ws As Worksheet, pt As PivotTable, c As Range
    Dim fails As Long, status As String

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws
    Application.CalculateFullRebuild

    For Each c In ThisWorkbook.Worksheets("Data_Quality").Range("D18:D26").Cells
        If Left$(CStr(c.Value), 4) = "FAIL" Then fails = fails + 1
    Next c
    status = IIf(fails = 0, "all reconciliations PASS", fails & " reconciliation(s) FAIL - see Data_Quality")

    ThisWorkbook.Worksheets("Dashboard").Range("B4").Value = _
        "Last refreshed: " & Format$(Now, "dd-mmm-yyyy hh:mm") & "  |  " & status

Tidy:
    Application.ScreenUpdating = True
    MsgBox "Refresh complete: " & status & ".", IIf(fails = 0, vbInformation, vbExclamation), "Commercial model"
    Exit Sub
Fail:
    status = "refresh error: " & Err.Description
    fails = -1
    Resume Tidy
End Sub
