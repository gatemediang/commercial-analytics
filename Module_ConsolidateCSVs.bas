Attribute VB_Name = "Module_ConsolidateCSVs"
Option Explicit

' Consolidates same-schema CSV files from a user-selected folder.
' The first CSV supplies the header row. Subsequent CSVs append data below it.

Public Sub ConsolidateCSVFiles()
    Dim folderPath As String, fileName As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim nextRow As Long, lastRow As Long, lastCol As Long
    Dim firstFile As Boolean

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing CSV files"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1) & Application.PathSeparator
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("CSV_Consolidated")
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.Name = "CSV_Consolidated"
    Else
        wsOut.Cells.Clear
    End If

    firstFile = True
    fileName = Dir(folderPath & "*.csv")

    Do While fileName <> ""
        Set wbSrc = Workbooks.Open(folderPath & fileName)
        Set wsSrc = wbSrc.Worksheets(1)

        lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

        If firstFile Then
            wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Copy _
                Destination:=wsOut.Cells(1, 1)
            firstFile = False
        ElseIf lastRow >= 2 Then
            nextRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row + 1
            wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy _
                Destination:=wsOut.Cells(nextRow, 1)
        End If

        wbSrc.Close SaveChanges:=False
        fileName = Dir()
    Loop

    wsOut.Columns.AutoFit
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "CSV consolidation complete: " & wsOut.Name, vbInformation
End Sub
