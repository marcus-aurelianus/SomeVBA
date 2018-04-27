Attribute VB_Name = "Module4"
Sub Mail_Selection_Range_Outlook_Body()
'Written by: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    'Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    Set rng = Sheets("DATA3").Range("A2:AO49").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set Sourcewb = ActiveWorkbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook

    ' Determine the Excel version, and file extension and format.
    With Destwb
        If Val(Application.Version) < 12 Then
            ' For Excel 2000-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            ' For Excel 2007-2010, exit the subroutine if you answer
            ' NO in the security dialog that is displayed when you copy
            ' a sheet from an .xlsm file with macros disabled.
            If Sourcewb.Name = .Name Then
                With Application
                    .ScreenUpdating = True
                    .EnableEvents = True
                End With
                MsgBox "You answered NO in the security dialog."
                Exit Sub
            Else
                Select Case Sourcewb.FileFormat
                Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
                Case 52:
                    If .HasVBProject Then
                        FileExtStr = ".xlsm": FileFormatNum = 52
                    Else
                        FileExtStr = ".xlsx": FileFormatNum = 51
                    End If
                Case 56: FileExtStr = ".xls": FileFormatNum = 56
                Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
                End Select
            End If
        End If
    End With
    With Destwb.Sheets(1).UsedRange
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False
    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Part of " & Sourcewb.Name & " " _
                 & Format(Now, "dd-mmm-yy h-mm-ss")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    


    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, _
                FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
             .To = "Sample@Sample.com"
             .cc = ""
             .BCC = ""
             .Subject = Format(Now, "dd-mmm-yy") & "  IE Report: Litho Scanner Loss Time - ALI4/ALD5_7"
             .HTMLBody = "<HTML><BODY>Dear all, <br><br> <p>Please find the report details for today below.</BODY></HTML>" & RangetoHTML(rng) & "<HTML><p>Regards<br>Marcus Li<br>IE & Planning<br>Ext 7445<p></HTML>"
             .Attachments.Add Destwb.FullName
             '.Attachments.Add ("Paths")
             .Send   'or use .Display
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With
    
    ' Delete the file after sending.
    Kill TempFilePath & TempFileName & FileExtStr

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function



