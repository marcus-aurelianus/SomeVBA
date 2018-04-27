Attribute VB_Name = "Module3"
Sub ReportOut()

Dim ewb, wb As Workbook 'e = external
Dim ws2, ews1, ews2, ews3 As Worksheet
Dim i As Integer

Set wb = ThisWorkbook
Set ewb = Workbooks.Open _
    ("Sample Path\02_Daily Report_Template v3.xlsx")
Set ws2 = wb.Sheets("TABLE2_NONEXPOSURE")
Set ews1 = ewb.Sheets("DATA1")
Set ews2 = ewb.Sheets("DATA2")
Set ews3 = ewb.Sheets("DATA3")

    ews1.Columns("A:AK").ClearContents
    ews2.Columns("A:D").ClearContents
    'ws2.Range("A1").Select
    'Selection.End(xlDown).Select '(This is the shift+end+down)
    Set r1 = ws2.Range(ws2.Cells(1, 1), ws2.Cells(WorksheetFunction.CountA(ws2.Columns(1)), 17))
    Set r2 = ws2.Range(ws2.Cells(1, 20), ws2.Cells(WorksheetFunction.CountA(ws2.Columns(1)), 40))
    Union(r1, r2).Copy
    ews1.Range("A1").PasteSpecial Paste:=xlPasteValues
    With ews1.PivotTables("PivotTable9").PivotFields("Final State(18)")
        .Orientation = xlRowField
        .Position = 1
    End With
    ews1.PivotTables("PivotTable9").PivotCache.Refresh
    ews1.Range("AS3:AV3000").Copy
    ews2.Range("A1").PasteSpecial Paste:=xlPasteValues
    ews2.PivotTables("PivotTable2").PivotCache.Refresh
    Sheets("DATA3").Select
    On Error Resume Next
    'ewb.SaveAs ("Sample path")
    ewb.SaveAs ("Sample Path\02_Daily Report_" & wb.Sheets("GUI").Range("C12") & ".xlsx")
    'MsgBox Mid(ws2.Cells(2, 1), 3, 2) & Mid(ws2.Cells(2, 1), 6, 2) & Right(ws2.Cells(2, 1), 2)
End Sub


