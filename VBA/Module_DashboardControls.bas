Attribute VB_Name = "Module_DashboardControls"
Option Explicit

' ---------------------------------------------------------------------------
' ResetDashboardFilters - back to All regions / categories / months, Base case
' AddDashboardButtons   - run once: draws clickable buttons to the right of
'                         the Dashboard and assigns the macros to them
' ---------------------------------------------------------------------------
Public Sub ResetDashboardFilters()
    ThisWorkbook.Names("selRegion").RefersToRange.Value = "All"
    ThisWorkbook.Names("selCategory").RefersToRange.Value = "All"
    ThisWorkbook.Names("selMonth").RefersToRange.Value = "All"
    ThisWorkbook.Names("selScenario").RefersToRange.Value = "Base"
End Sub

Public Sub AddDashboardButtons()
    Dim ws As Worksheet, shp As Shape, i As Long
    Dim captions As Variant, macros As Variant
    Dim leftPos As Double, topPos As Double

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i

    captions = Array("Refresh model", "Reset filters", "Export PDF", "Regional pack", "Run all scenarios", "Validate raw data")
    macros = Array("RefreshCommercialModel", "ResetDashboardFilters", "ExportDashboardToPDF", _
                   "ExportRegionalReportPack", "RunAllScenarios", "CleanseProcurementData")

    leftPos = ws.Range("R6").Left + 6
    topPos = ws.Range("R6").Top
    For i = LBound(captions) To UBound(captions)
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + i * 34, 130, 28)
        shp.Name = "btn_" & macros(i)
        shp.OnAction = macros(i)
        shp.Fill.ForeColor.RGB = RGB(31, 78, 120)
        shp.Line.Visible = msoFalse
        With shp.TextFrame2
            .TextRange.Text = captions(i)
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    Next i
    MsgBox "Buttons added to the Dashboard (column R).", vbInformation
End Sub
