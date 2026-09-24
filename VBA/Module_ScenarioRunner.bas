Attribute VB_Name = "Module_ScenarioRunner"
Option Explicit

' ---------------------------------------------------------------------------
' RunAllScenarios
' Pushes each scenario through Financial_Model and logs the outputs to
' Scenario_Log, then restores the scenario the user had selected. Saves
' switching the drop-down by hand and copying numbers into a board paper.
' ---------------------------------------------------------------------------
Public Sub RunAllScenarios()
    Dim ws As Worksheet, fm As Worksheet, sc As Range
    Dim original As String, r As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set fm = ThisWorkbook.Worksheets("Financial_Model")
    Set ws = GetOrCreateSheet("Scenario_Log")
    ws.Cells.Clear
    ws.Range("A1:H1").Value = Array("Scenario", "Revenue", "Total Opex", "EBITDA", "EBITDA margin", _
                                    "Procurement saving", "EBITDA vs 2025", "Run time")
    ws.Range("A1:H1").Font.Bold = True

    original = CStr(ThisWorkbook.Names("selScenario").RefersToRange.Value)
    r = 2
    For Each sc In ThisWorkbook.Names("lstScenario").RefersToRange.Cells
        ThisWorkbook.Names("selScenario").RefersToRange.Value = sc.Value
        Application.Calculate
        ws.Cells(r, 1).Value = sc.Value
        ws.Cells(r, 2).Value = fm.Range("C16").Value
        ws.Cells(r, 3).Value = fm.Range("C22").Value
        ws.Cells(r, 4).Value = fm.Range("C23").Value
        ws.Cells(r, 5).Value = fm.Range("C24").Value
        ws.Cells(r, 6).Value = -fm.Range("D17").Value
        ws.Cells(r, 7).Value = fm.Range("D23").Value
        ws.Cells(r, 8).Value = Now
        r = r + 1
    Next sc

    ws.Range("B2:D" & r - 1 & ",F2:G" & r - 1).NumberFormat = "£#,##0;(£#,##0)"
    ws.Range("E2:E" & r - 1).NumberFormat = "0.0%"
    ws.Range("H2:H" & r - 1).NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Columns("A:H").AutoFit

Tidy:
    ThisWorkbook.Names("selScenario").RefersToRange.Value = original
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox "Scenario_Log updated for " & (r - 2) & " scenarios.", vbInformation
    Exit Sub
Fail:
    MsgBox "Scenario run stopped: " & Err.Description, vbCritical
    Resume Tidy
End Sub
