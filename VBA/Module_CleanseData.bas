Attribute VB_Name = "Module_CleanseData"
Option Explicit

' ---------------------------------------------------------------------------
' CleanseProcurementData
' Validates every row of tblProcurement (sheet Procurement_Raw) and:
'   1. writes a flag to the DataQualityFlag column,
'   2. writes every exception to a timestamped DQ_Log sheet.
' Source text (supplier, site, category) is NOT overwritten; standardisation
' happens in the Cleaned_Procurement formula layer. The flag hierarchy matches
' Cleaned_Procurement!N so the macro and the formulas can be cross-checked.
' Arrays + Scripting.Dictionary keep it fast (O(n), not COUNTIF per row).
' ---------------------------------------------------------------------------
Public Sub CleanseProcurementData()
    Dim lo As ListObject, loSup As ListObject
    Dim data As Variant, supData As Variant, flags() As Variant, logArr() As Variant
    Dim cSup As Long, cSite As Long, cCat As Long, cInv As Long, cSpend As Long, cFlag As Long
    Dim i As Long, n As Long, nEx As Long
    Dim idCount As Object, seenFull As Object, supMaster As Object
    Dim supplier As String, site As String, inv As String, cat As String, fullKey As String, dq As String
    Dim spend As Variant
    Dim wsLog As Worksheet
    Dim tStart As Double

    tStart = Timer
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set lo = ThisWorkbook.Worksheets("Procurement_Raw").ListObjects("tblProcurement")
    If lo.DataBodyRange Is Nothing Then
        MsgBox "tblProcurement has no rows.", vbExclamation
        GoTo Tidy
    End If

    cSup = lo.ListColumns("Supplier").Index
    cSite = lo.ListColumns("SiteID").Index
    cCat = lo.ListColumns("Category").Index
    cInv = lo.ListColumns("InvoiceID").Index
    cSpend = lo.ListColumns("GrossSpend").Index
    cFlag = lo.ListColumns("DataQualityFlag").Index

    data = lo.DataBodyRange.Value
    n = UBound(data, 1)
    ReDim flags(1 To n, 1 To 1)
    ReDim logArr(1 To n, 1 To 7)

    ' supplier master: supplier -> category (case-insensitive)
    Set supMaster = CreateObject("Scripting.Dictionary")
    supMaster.CompareMode = vbTextCompare
    Set loSup = ThisWorkbook.Worksheets("Supplier_Lookup").ListObjects("tblSuppliers")
    supData = loSup.DataBodyRange.Value
    For i = 1 To UBound(supData, 1)
        If Len(Trim$(CStr(supData(i, 1)))) > 0 Then
            If Not supMaster.Exists(Trim$(CStr(supData(i, 1)))) Then
                supMaster.Add Trim$(CStr(supData(i, 1))), NormalizeCategory(CStr(supData(i, 2)))
            End If
        End If
    Next i

    ' pass 1: how many times does each invoice ID appear?
    Set idCount = CreateObject("Scripting.Dictionary")
    idCount.CompareMode = vbTextCompare
    For i = 1 To n
        inv = Trim$(CStr(data(i, cInv)))
        If Len(inv) > 0 Then idCount(inv) = idCount(inv) + 1
    Next i

    ' pass 2: classify each row (first failing rule wins)
    Set seenFull = CreateObject("Scripting.Dictionary")
    seenFull.CompareMode = vbTextCompare
    For i = 1 To n
        supplier = Trim$(CStr(data(i, cSup)))
        site = Trim$(CStr(data(i, cSite)))
        inv = Trim$(CStr(data(i, cInv)))
        cat = NormalizeCategory(CStr(data(i, cCat)))
        spend = data(i, cSpend)

        dq = "OK"
        If supplier = "" Then
            dq = "Missing supplier"
        ElseIf site = "" Then
            dq = "Missing site"
        ElseIf inv = "" Then
            dq = "Missing invoice"
        ElseIf IsEmpty(spend) Or Not IsNumeric(spend) Then
            dq = "Missing spend"
        Else
            fullKey = inv & "|" & supplier & "|" & Format$(CDbl(spend), "0.00")
            If seenFull.Exists(fullKey) Then
                dq = "Duplicate invoice"                 ' true duplicate: exclude 2nd copy
            ElseIf idCount(inv) > 1 Then
                dq = "Duplicate ID - review"             ' same ID, different supplier/amount
            ElseIf supMaster.Exists(supplier) Then
                If StrComp(supMaster(supplier), cat, vbTextCompare) <> 0 Then dq = "Category mismatch"
            End If
            If Not seenFull.Exists(fullKey) Then seenFull.Add fullKey, i
        End If

        flags(i, 1) = dq
        If dq <> "OK" Then
            nEx = nEx + 1
            logArr(nEx, 1) = Now
            logArr(nEx, 2) = i + lo.HeaderRowRange.Row      ' worksheet row
            logArr(nEx, 3) = inv
            logArr(nEx, 4) = supplier
            logArr(nEx, 5) = site
            logArr(nEx, 6) = spend
            logArr(nEx, 7) = dq
        End If
    Next i

    lo.ListColumns(cFlag).DataBodyRange.Value = flags

    ' exception log
    Set wsLog = GetOrCreateSheet("DQ_Log")
    wsLog.Cells.Clear
    wsLog.Range("A1:G1").Value = Array("RunTime", "SheetRow", "InvoiceID", "Supplier", "SiteID", "GrossSpend", "Flag")
    wsLog.Range("A1:G1").Font.Bold = True
    If nEx > 0 Then
        wsLog.Range("A2").Resize(nEx, 7).Value = SliceRows(logArr, nEx, 7)
        wsLog.Columns("A").NumberFormat = "dd-mmm-yyyy hh:mm"
        wsLog.Columns("F").NumberFormat = "£#,##0.00"
    End If
    wsLog.Columns("A:G").AutoFit

Tidy:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If n > 0 Then
        MsgBox "Validated " & Format$(n, "#,##0") & " invoices in " & Format$(Timer - tStart, "0.0") & "s." & vbCrLf & _
               nEx & " exception(s) written to DQ_Log. Source text was not changed.", vbInformation, "Procurement validation"
    End If
    Exit Sub

Fail:
    MsgBox "Validation stopped: " & Err.Description, vbCritical
    Resume Tidy
End Sub

Private Function NormalizeCategory(ByVal rawCategory As String) As String
    Select Case LCase$(Trim$(rawCategory))
        Case "medical supplies": NormalizeCategory = "Medical Supplies"
        Case "food & catering": NormalizeCategory = "Food & Catering"
        Case "facilities": NormalizeCategory = "Facilities"
        Case "utilities": NormalizeCategory = "Utilities"
        Case "it & telecoms": NormalizeCategory = "IT & Telecoms"
        Case "vehicle & fuel": NormalizeCategory = "Vehicle & Fuel"
        Case "agency labour": NormalizeCategory = "Agency Labour"
        Case "office & training": NormalizeCategory = "Office & Training"
        Case Else: NormalizeCategory = "Unmapped"
    End Select
End Function

Private Function SliceRows(ByVal src As Variant, ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim out() As Variant, r As Long, c As Long
    ReDim out(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            out(r, c) = src(r, c)
        Next c
    Next r
    SliceRows = out
End Function

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
