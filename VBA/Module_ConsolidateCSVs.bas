Attribute VB_Name = "Module_ConsolidateCSVs"
Option Explicit

' ---------------------------------------------------------------------------
' ConsolidateCSVFiles
' Appends every same-schema CSV in a chosen folder into CSV_Consolidated.
'  - the first CSV sets the expected header
'  - any file whose header differs is skipped and reported, not guessed at
'  - a SourceFile column records where each row came from (audit trail)
'  - output is converted to an Excel Table (tblConsolidated)
' Test with the Sample_Procurement_Files folder supplied with the project.
' ---------------------------------------------------------------------------
Public Sub ConsolidateCSVFiles()
    Dim folderPath As String, fileName As String
    Dim wbSrc As Workbook, wsSrc As Worksheet, wsOut As Worksheet
    Dim expectedHeader As String, thisHeader As String
    Dim data As Variant, lastRow As Long, lastCol As Long, nextRow As Long
    Dim filesLoaded As Long, rowsLoaded As Long, skipped As String
    Dim lo As ListObject

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder that holds the CSV extracts"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1) & Application.PathSeparator
    End With

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsOut = GetOrCreateSheet("CSV_Consolidated")
    Do While wsOut.ListObjects.Count > 0
        wsOut.ListObjects(1).Unlist
    Loop
    wsOut.Cells.Clear

    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        Set wbSrc = Workbooks.Open(Filename:=folderPath & fileName, ReadOnly:=True)
        Set wsSrc = wbSrc.Worksheets(1)
        lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
        thisHeader = HeaderSignature(wsSrc, lastCol)

        If expectedHeader = "" Then
            expectedHeader = thisHeader
            wsOut.Range("A1").Resize(1, lastCol).Value = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Value
            wsOut.Cells(1, lastCol + 1).Value = "SourceFile"
            nextRow = 2
        End If

        If thisHeader <> expectedHeader Then
            skipped = skipped & vbCrLf & "  - " & fileName & " (header mismatch)"
        ElseIf lastRow >= 2 Then
            data = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Value
            wsOut.Cells(nextRow, 1).Resize(lastRow - 1, lastCol).Value = data
            wsOut.Cells(nextRow, lastCol + 1).Resize(lastRow - 1, 1).Value = fileName
            nextRow = nextRow + lastRow - 1
            rowsLoaded = rowsLoaded + lastRow - 1
            filesLoaded = filesLoaded + 1
        End If

        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
        fileName = Dir()
    Loop

    If rowsLoaded > 0 Then
        Set lo = wsOut.ListObjects.Add(xlSrcRange, wsOut.Range("A1").CurrentRegion, , xlYes)
        lo.Name = "tblConsolidated"
        wsOut.Columns.AutoFit
    End If

Tidy:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox filesLoaded & " file(s) and " & Format$(rowsLoaded, "#,##0") & " row(s) loaded into CSV_Consolidated." & _
           IIf(Len(skipped) > 0, vbCrLf & vbCrLf & "Skipped:" & skipped, ""), vbInformation, "CSV consolidation"
    Exit Sub

Fail:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    MsgBox "Consolidation stopped on " & fileName & ": " & Err.Description, vbCritical
    Resume Tidy
End Sub

Private Function HeaderSignature(ByVal ws As Worksheet, ByVal lastCol As Long) As String
    Dim c As Long, s As String
    For c = 1 To lastCol
        s = s & "|" & LCase$(Trim$(CStr(ws.Cells(1, c).Value)))
    Next c
    HeaderSignature = s
End Function
