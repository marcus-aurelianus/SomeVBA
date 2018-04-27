VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Data_Refresh 
   Caption         =   "Data_Refresh"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3420
   OleObjectBlob   =   "Data_Refresh.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Data_Refresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Confirm_Click()

Dim Start_Datee, End_Datee, size, i, x As Integer
Dim inputdate, inputdate1, cmdtxt As String
Dim cx() As String

cmdtxt = ""
 Start_Datee = Start_Date.ListIndex
 End_Datee = End_Date.ListIndex
 
 If Start_Datee < End_Datee Then
    MsgBox "[Start Date] cannot be later then [End Date]!"
 Else
 
 Data_Refresh.Hide
 'MsgBox "aa"
Select Case Start_Datee + End_Datee
    Case 0 'today
    x = 100
todaytoo:

    inputdate = (Right(Date, 2)) & "-" & Mid(Format(Date, "yyyy/mmm/dd"), 6, 3) & "-" & Left(Format(Date, "yyyy/mmm/dd"), 4)
    inputdate1 = (Right(Date - 1, 2)) & "-" & Mid(Format(Date - 1, "yyyy/mmm/dd"), 6, 3) & "-" & Left(Format(Date - 1, "yyyy/mmm/dd"), 4)
    
    cmdtxt = cmdtxt & _
    "SELECT a.histtype,a.histdate,a.flag,a.eqpid,a.starttime,a.endtime,a.duration,a.lotid,a.qty,a.text1,a.text2,a.text3,a.rectime,a.s_status,a.s_eqs_1,a.s_eqs_2,a.insert_by,nvl(b.EQPRUNRECIPE ,0) eqprunrecipe, max(b.TRACKOUTTIME) tractouttime,a.new_rnum_ne " & _
    "FROM ismrt.ismrt_tbl_stg_raw_eqplog_ne a left join (select eqpid,LOTID, EQPRUNRECIPE,startruntime,endruntime,TO_Char(TRACKOUTTIME,'DD-MON-YYYY HH24:MI:SS') TRACKOUTTIME from ismrt.ISMRT_MV_PLS_TBL_LOGSHEET_FULL)b on a.lotid = b.lotid and substr(a.eqpid,0,6) = b.eqpid and(substr(a.starttime,0,6) = substr(b.startruntime,0,6) or substr(a.endtime,0,6) = substr(b.endruntime,0,6))" & _
    "where endtime >= to_date('" & inputdate1 & " 08:00:00','DD-MON-YYYY HH24:MI:SS') " & _
    "AND endtime <= to_date('" & inputdate & " 08:00:00','DD-MON-YYYY HH24:MI:SS') " & _
    "AND histdate=to_date('" & inputdate1 & " 00:00:00','DD-MON-YYYY HH24:MI:SS') "
    cmdtxt = cmdtxt & "group by a.histtype,a.histdate,a.flag,a.eqpid,a.starttime,a.endtime,a.duration,a.lotid,a.qty,a.text1,a.text2 , a.text3, a.rectime, a.s_status, a.s_eqs_1, a.s_eqs_2, a.insert_by, nvl(b.EQPRUNRECIPE, 0), a.new_rnum_ne "
    
    If x <> 100 Then GoTo histback:
    cmdtxt = cmdtxt & "ORDER BY EQPID ASC, NEW_RNUM_NE ASC"
    
    ThisWorkbook.Connections("Today").OLEDBConnection.CommandText = cmdtxt
    
    ThisWorkbook.Connections("Today").OLEDBConnection.Connection = "OLEDB;DSN=mes;UID=iep;PWD=iep"
    ThisWorkbook.Connections("Today").OLEDBConnection.Refresh
    ThisWorkbook.Connections("Today").OLEDBConnection.Connection = "OLEDB;DSN=mes;UID=iep;"
    Sheets("TABLE2_NONEXPOSURE").Cells().ClearContents
    Sheets("TABLE2_NONEXPOSURE").Cells().Interior.ColorIndex = 0
    Sheets("TABLE2_NONEXPOSURE").Cells().Font.Bold = False
    Sheets("Today").Range(Sheets("Today").Cells(1, 1), Sheets("Today").Cells(WorksheetFunction.CountA(Sheets("Today").Columns(1)), WorksheetFunction.CountA(Sheets("Today").Rows(1)) - 1)).Copy
    Sheets("TABLE2_NONEXPOSURE").Cells().PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Sheets("GUI").Range("C12").Value = inputdate1

    Case Else 'hist
     size = WorksheetFunction.CountA(ThisWorkbook.Sheets("Hist").Columns(1))
     ReDim cx(size)
     x = 3
     For i = 2 To size
     cx(i) = Left(Sheets("Hist").Cells(x, 1), 10)
     x = x + 1
     Next i
     'MsgBox "a"
    If End_Datee = 0 Then
        inputdate = (Right(Date, 2)) & "-" & Mid(Format(Date, "yyyy/mmm/dd"), 6, 3) & "-" & Left(Format(Date, "yyyy/mmm/dd"), 4)
    Else
        inputdate = (Right(CDate(cx(End_Datee + 1)) + 1, 2)) & "-" & Mid(Format(CDate(cx(End_Datee + 1)) + 1, "yyyy/mmm/dd"), 6, 3) & "-" & Left(CDate(cx(End_Datee + 1)) + 1, 4)
    End If
    inputdate1 = (Right(CDate(cx(Start_Datee + 1)), 2)) & "-" & Mid(Format(CDate(cx(Start_Datee + 1)), "yyyy/mmm/dd"), 6, 3) & "-" & Left(CDate(cx(Start_Datee + 1)), 4)
    

    cmdtxt = "SELECT a.histtype,a.histdate,a.flag,a.eqpid,a.starttime,a.endtime,a.duration,a.lotid,a.qty,a.text1,a.text2,a.text3,a.rectime,a.s_status,a.s_eqs_1,a.s_eqs_2,a.insert_by,nvl(b.EQPRUNRECIPE ,0) eqprunrecipe, max(b.TRACKOUTTIME) tractouttime,a.new_rnum_ne " & _
    "FROM ismrt.ismrt_tbl_stg_raw_eqplognehist a left join (select eqpid,LOTID, EQPRUNRECIPE,startruntime,endruntime,TO_Char(TRACKOUTTIME,'DD-MON-YYYY HH24:MI:SS') TRACKOUTTIME from ismrt.ISMRT_MV_PLS_TBL_LOGSHEET_FULL) b on a.lotid = b.lotid And substr(a.eqpid, 0, 6) = b.eqpid And (substr(a.starttime, 0, 6) = substr(b.startruntime, 0, 6) Or substr(a.endtime, 0, 6) = substr(b.endruntime, 0, 6))" & _
    "where endtime >= to_date('" & inputdate1 & " 08:00:00','DD-MON-YYYY HH24:MI:SS') " & _
    "AND endtime <= to_date('" & inputdate & " 08:00:00','DD-MON-YYYY HH24:MI:SS') " & _
    "AND (histdate=to_date('" & inputdate1 & " 00:00:00','DD-MON-YYYY HH24:MI:SS') "

    For x = -1 To Start_Datee - 2
    inputdate = (Right(CDate(cx(Start_Datee + 1 - x)) + 1, 2)) & "-" & Mid(Format(CDate(cx(Start_Datee + 1 - x)) + 1, "yyyy/mmm/dd"), 6, 3) & "-" & Left(CDate(cx(Start_Datee + 1 - x)) + 1, 4)
    cmdtxt = cmdtxt & "or histdate=to_date('" & inputdate & " 00:00:00','DD-MON-YYYY HH24:MI:SS') "
    Next x
    
    cmdtxt = cmdtxt & ") group by a.histtype,a.histdate,a.flag,a.eqpid,a.starttime,a.endtime,a.duration,a.lotid,a.qty,a.text1,a.text2,a.text3,a.rectime,a.s_status,a.s_eqs_1,a.s_eqs_2,a.insert_by,nvl(b.EQPRUNRECIPE ,0),a.new_rnum_ne "
    
    If End_Datee = 0 Then
        cmdtxt = cmdtxt & " union all ("
        'MsgBox cmdtxt
        GoTo todaytoo:
    End If
histback:
    
    If End_Datee = 0 Then
    cmdtxt = cmdtxt & ")ORDER BY HISTDATE asc, EQPID ASC, NEW_RNUM_NE ASC"
    Sheets("GUI").Range("C12").Value = inputdate1 & "_to_" & inputdate
    Else
    cmdtxt = cmdtxt & "ORDER BY HISTDATE asc, EQPID ASC, NEW_RNUM_NE ASC"
    Sheets("GUI").Range("C12").Value = inputdate1
    End If
    
    ThisWorkbook.Connections("Hist").OLEDBConnection.CommandText = cmdtxt
    
    ThisWorkbook.Connections("Hist").OLEDBConnection.Connection = "OLEDB;DSN=mes;UID=iep;PWD=iep"
    ThisWorkbook.Connections("Hist").Refresh
    ThisWorkbook.Connections("Hist").OLEDBConnection.Connection = "OLEDB;DSN=mes;UID=iep;"
    Sheets("TABLE2_NONEXPOSURE").Cells().ClearContents
    Sheets("TABLE2_NONEXPOSURE").Cells().Interior.ColorIndex = 0
    Sheets("TABLE2_NONEXPOSURE").Cells().Font.Bold = False
    Sheets("Hist").Range(Sheets("Hist").Cells(1, 1), Sheets("Hist").Cells(WorksheetFunction.CountA(Sheets("Hist").Columns(1)), WorksheetFunction.CountA(Sheets("Hist").Rows(1)) - 1)).Copy
    Sheets("TABLE2_NONEXPOSURE").Cells().PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Select

End If

End Sub

Private Sub Start_Date_Change()

End Sub

Private Sub UserForm_Initialize()

Dim x, size As Integer

x = 2
With Start_Date
     Do While Sheets("Hist").Cells(x, 1) <> ""
        .AddItem Left(Sheets("Hist").Cells(x, 1), 10)
        x = x + 1
     Loop
     .text = Left(Sheets("Hist").Cells(2, 1), 10)
End With

x = 2
With End_Date
     Do While Sheets("Hist").Cells(x, 1) <> ""
        .AddItem Left(Sheets("Hist").Cells(x, 1), 10)
        x = x + 1
     Loop
     .text = Left(Sheets("Hist").Cells(2, 1), 10)
End With


End Sub
