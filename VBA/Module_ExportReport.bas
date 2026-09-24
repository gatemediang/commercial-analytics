Attribute VB_Name = "Module_ExportReport"
Option Explicit

' ---------------------------------------------------------------------------
' ExportDashboardToPDF      - current view of the Dashboard to one PDF
' ExportRegionalReportPack  - one PDF per region (plus group) in a chosen folder,
'                             restoring the user's selection afterwards
' Exported PDFs contain aggregated figures only (no personal data).
' ---------------------------------------------------------------------------
Public Sub ExportDashboardToPDF()
    Dim target As Variant
    target = Application.GetSaveAsFilename( _
        InitialFileName:="Commercial_Dashboard_" & Format$(Date, "yyyymmdd") & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")
    If VarType(target) = vbBoolean Then Exit Sub
    PrepareDashboardPage
    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(target), _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

Public Sub ExportRegionalReportPack()
    Dim folderPath As String, original As String, region As Variant, fileOut As String
    Dim regions As Range, n As Long

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose a folder for the regional report pack"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1) & Application.PathSeparator
    End With

    On Error GoTo Fail
    Application.ScreenUpdating = False
    original = CStr(ThisWorkbook.Names("selRegion").RefersToRange.Value)
    Set regions = ThisWorkbook.Names("lstRegion").RefersToRange
    PrepareDashboardPage

    For Each region In regions.Cells
        ThisWorkbook.Names("selRegion").RefersToRange.Value = region.Value
        Application.Calculate
        fileOut = folderPath & "Commercial_Dashboard_" & IIf(region.Value = "All", "Group", region.Value) & "_" & Format$(Date, "yyyymmdd") & ".pdf"
        ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileOut, _
            Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
        n = n + 1
    Next region

Tidy:
    ThisWorkbook.Names("selRegion").RefersToRange.Value = original
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox n & " PDF(s) saved to " & folderPath, vbInformation, "Regional report pack"
    Exit Sub
Fail:
    MsgBox "Report pack stopped: " & Err.Description, vbCritical
    Resume Tidy
End Sub

Private Sub PrepareDashboardPage()
    With ThisWorkbook.Worksheets("Dashboard").PageSetup
        .PrintArea = "$A$1:$Q$96"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Synthetic data - portfolio project - &D"
    End With
End Sub
