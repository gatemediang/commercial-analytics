Attribute VB_Name = "Module_CleanseData"
Option Explicit

' Cleans and validates the synthetic procurement dataset.
' Designed for the Excel table named tblProcurement on Procurement_Raw.

Public Sub CleanseProcurementData()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As ListRow
    Dim supplier As String, siteID As String, category As String, invoiceID As String
    Dim spendValue As Variant
    Dim dq As String

    Set ws = ThisWorkbook.Worksheets("Procurement_Raw")
    Set lo = ws.ListObjects("tblProcurement")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each r In lo.ListRows
        supplier = Trim(CStr(r.Range.Cells(1, lo.ListColumns("Supplier").Index).Value))
        siteID = Trim(CStr(r.Range.Cells(1, lo.ListColumns("SiteID").Index).Value))
        category = NormalizeCategory(CStr(r.Range.Cells(1, lo.ListColumns("Category").Index).Value))
        invoiceID = Trim(CStr(r.Range.Cells(1, lo.ListColumns("InvoiceID").Index).Value))
        spendValue = r.Range.Cells(1, lo.ListColumns("GrossSpend").Index).Value

        r.Range.Cells(1, lo.ListColumns("Supplier").Index).Value = supplier
        r.Range.Cells(1, lo.ListColumns("SiteID").Index).Value = siteID
        r.Range.Cells(1, lo.ListColumns("Category").Index).Value = category

        dq = "OK"
        If supplier = "" Then dq = "Missing supplier"
        If siteID = "" Then dq = "Missing site"
        If invoiceID = "" Then dq = "Missing invoice"
        If IsEmpty(spendValue) Or spendValue = "" Then dq = "Missing spend"

        If invoiceID <> "" Then
            If Application.WorksheetFunction.CountIf(lo.ListColumns("InvoiceID").DataBodyRange, invoiceID) > 1 Then
                dq = "Duplicate invoice"
            End If
        End If

        r.Range.Cells(1, lo.ListColumns("DataQualityFlag").Index).Value = dq
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Procurement cleansing and validation complete.", vbInformation
End Sub

Private Function NormalizeCategory(ByVal rawCategory As String) As String
    Select Case LCase(Trim(rawCategory))
        Case "medical supplies": NormalizeCategory = "Medical Supplies"
        Case "food & catering": NormalizeCategory = "Food & Catering"
        Case "facilities": NormalizeCategory = "Facilities"
        Case "utilities": NormalizeCategory = "Utilities"
        Case "it & telecoms": NormalizeCategory = "IT & Telecoms"
        Case "vehicle & fuel": NormalizeCategory = "Vehicle & Fuel"
        Case "agency labour": NormalizeCategory = "Agency Labour"
        Case "office & training": NormalizeCategory = "Office & Training"
        Case Else: NormalizeCategory = Application.WorksheetFunction.Proper(Trim(rawCategory))
    End Select
End Function
