Attribute VB_Name = "Module_ExportReport"
Option Explicit

' Exports the Dashboard worksheet to a PDF selected by the user.

Public Sub ExportDashboardToPDF()
    Dim ws As Worksheet
    Dim targetFile As Variant

    Set ws = ThisWorkbook.Worksheets("Dashboard")

    targetFile = Application.GetSaveAsFilename( _
        InitialFileName:="Commercial_Dashboard_" & Format(Date, "yyyymmdd") & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If targetFile = False Then Exit Sub

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(targetFile), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Dashboard exported to PDF.", vbInformation
End Sub
