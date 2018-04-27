Attribute VB_Name = "Module2"
Sub Analyze_Data()

Dim i, cc As Long
Dim a, b As Long
Dim boss As Long
Dim x, x1, y, z, x2, x3 As Long
Dim text, text1 As String
Dim count As Integer
Dim total As Long
Dim wb As Workbook
Dim ws2, wst, wsr, wsp As Worksheet 'wst = temp, wsr = resist, wsp = position
Dim tbrng As Integer

ThisWorkbook.Save

Set wb = ThisWorkbook
Set ws2 = Sheet1
Set wst = Sheet6
Set wsr = Sheet7
Set wsp = Sheet8

Application.DisplayAlerts = False
'Private Sub Count_Total()

If ws2.Cells(1, WorksheetFunction.CountA(ws2.Rows(1))) = "Classification" Then ws2.Range(ws2.Columns(WorksheetFunction.CountA(ws2.Rows(1)) - 27), ws2.Columns(WorksheetFunction.CountA(ws2.Rows(1)))).ClearContents

tbrng = WorksheetFunction.CountA(ws2.Rows(1))
total = WorksheetFunction.CountA(ws2.Columns(1))

'GoTo db:

ws2.Cells(1, tbrng + 2) = "Dummy Row"
'Private Sub Final State()
'=IF(N2="Exposure","Run",IF(P2="Available","Loss",O2))
        x = 2
        'ws2.Range("R2:R999999").ClearContents
        ws2.Cells(1, tbrng + 1).Value = "Final State(18)"
        Do While x < total
            If Left(ws2.Cells(x, 14), 8) = "Exposure" Then
                ws2.Cells(x, tbrng + 1).Value = "Run"
                Else
                If Left(ws2.Cells(x, 16), 9) = "Available" Then
                    ws2.Cells(x, tbrng + 1).Value = "Loss"
                    Else
                    ws2.Cells(x, tbrng + 1).Value = ws2.Cells(x, 15)
                End If
            End If
            x = x + 1
        Loop

'Private Sub Logic1_Click()
        i = 2
        x = 2
        x1 = 0
        'ws2.Range("T2:T999999").ClearContents
        ws2.Cells(1, tbrng + 3).Value = "Logic1" & Chr(13) & "(Test_Loss)"
        Do While i < total
            If Left(ws2.Cells(x, tbrng + 1), 4) = "Loss" And Left(ws2.Cells(x - 1, tbrng + 1), 4) <> "Loss" Then x1 = x
            
            If ws2.Cells(x, 15) = "TEST" And Left(ws2.Cells(x, tbrng + 1), 4) = "Loss" Then
                Do While Left(ws2.Cells(x, tbrng + 1), 4) = "Loss"
                    x = x + 1
                Loop
            
                Do While x <> x1
                    ws2.Cells(x1, tbrng + 3).Value = "TEST_Loss"
                    x1 = x1 + 1
                Loop
                x = x - 1
            End If
            i = i + 1
            x = x + 1
        Loop

'Private Sub Logic2_Click()
        x = 2
        i = 2
        'ws2.Range("U2:U999999").ClearContents
         ws2.Cells(1, tbrng + 4).Value = "Logic2" & Chr(13) & "(Cold Start)"
        Do While i < total
            If Left(ws2.Cells(x, tbrng + 1), 4) = "Loss" And Left(ws2.Cells(x - 1, 16), 13) = "Non-Available" Then
                Do While Left(ws2.Cells(x, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x, tbrng + 4).Value = "Cold Start"
                    x = x + 1
                    i = i + 1
                Loop
            Else
            x = x + 1
            i = i + 1
            End If
        Loop

'Private Sub IRIS_No_Reticle_Click()
        i = 2
        x = 2
        boss = 0
        z = 0
        cc = 0
        x2 = 0
        'ws2.Range("X2:Y999999").ClearContents
        ws2.Cells(1, tbrng + 7).Value = "IRIS"
        ws2.Cells(1, tbrng + 8).Value = "No Reticle"
        Do While i < total
            If ws2.Cells(x, tbrng + 1) = "Loss" Then
                text = ws2.Cells(x, 12)
                If Mid(text, 2, 23) = "Reticles for next batch" Or Mid(text, 3, 23) = "Reticles for next batch" And ws2.Cells(x, tbrng + 1) = "Loss" Then
                    x1 = x
                    Do While 1
                        If ws2.Cells(x1, tbrng + 1) = "Loss" Then
                            For a = 1 To Len(ws2.Cells(x1, 12)) + 1
                                If Mid(ws2.Cells(x1, 12), a, 6) = "inspec" Or boss = 12 Then
                                    If boss = 0 Then x2 = x1
                                    boss = 12
                                    z = z + 1
                                    GoTo dd:
                                End If
                            Next a
dd:
                        Else
                            Do While cc <= (x - x1) And z <> 0
                            ws2.Cells(x - cc + 1, tbrng + 7).Value = "IRIS"
                            ws2.Cells(x - cc + 1, tbrng + 8).Value = ""
                            If ws2.Cells(x + 1, tbrng + 1) <> "Loss" Then ws2.Cells(x + 1, tbrng + 7).Value = ""
                            cc = cc + 1
                            Loop
                            GoTo ee:
                        End If
                        If ws2.Cells(x1, tbrng + 7) = "" Then
                            ws2.Cells(x1, tbrng + 8).Value = "No Reticle"
                            If ws2.Cells(x + 1, tbrng + 1) = "Loss" Then ws2.Cells(x + 1, tbrng + 8).Value = "No Reticle"
                        End If
                        x1 = x1 - 1
                    Loop
                End If
            End If
ee:
            x = x + 1
            i = i + 1
            boss = 0
            z = 0
            cc = 0
        Loop

'Private Sub IQC_Click()
        x = 2
        z = 0
        x1 = 0
        x2 = 1
        'ws2.Range("Z2:Z999999").ClearContents
        ws2.Cells(1, tbrng + 9) = "IQC"
        Do While x < total
            If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: EC-0060" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 9).Value = "IQC"
                    x1 = x1 - 1
                Loop
            Else
                If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: EC-0060" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 27) <> "DEACTIVATE WARNING: EC-0060"
                        ws2.Cells(x, tbrng + 9).Value = "IQC"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 14) = "ACTIVATE WARNING: EC-0060" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 9).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 9).Value = "IQC"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 9).Value = "IQC"
                                x1 = x1 + 1
                            Loop
                            GoTo cc:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: EC-0060" Then ws2.Cells(x, tbrng + 9).Value = "IQC"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 9).Value = "IQC"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x = x + 1
            z = 0
            x2 = 1
cc:
        Loop
'Private Sub Test_Queue_n_IQC_Click()
        x = 2
        z = 0
        boss = 0
        x1 = 0
        x2 = 1
        'ws2.Range("W2:W999999").ClearContents
        'ws2.Range("AA2:AA999999").ClearContents
        ws2.Cells(1, tbrng + 6).Value = "Test Queue_IQC"
        ws2.Cells(1, tbrng + 10).Value = "Test Queue"
        Do While x < total
            If Left(ws2.Cells(x, 12), 9) = "Ran Queue" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    If Left(ws2.Cells(x, 12), 13) = "Ran Queue IQC" Or boss = 1 Then
                        ws2.Cells(x1, tbrng + 6).Value = "Test Queue_IQC"
                        boss = 1
                    Else
                        ws2.Cells(x1, tbrng + 10).Value = "Test Queue"
                    End If
                    x1 = x1 - 1
                Loop
            Else
                If Left(ws2.Cells(x, 12), 13) = "Running Queue" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 12), 9) <> "Ran Queue"
                        If Left(ws2.Cells(x, 12), 17) = "Running Queue IQC" Or boss = 1 Then
                            ws2.Cells(x, tbrng + 6).Value = "Test Queue_IQC"
                            boss = 1
                        Else
                            ws2.Cells(x, tbrng + 10).Value = "Test Queue"
                        End If
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 12), 13) = "Running Queue" Then
                            If Left(ws2.Cells(x, 12), 17) <> "Running Queue IQC" Or boss <> 1 Then
                                Do While z > 0
                                    ws2.Cells(x - z, tbrng + 10).Value = ""
                                    If x1 = x - z Then ws2.Cells(x1, tbrng + 10).Value = "Test Queue"
                                    z = z - 1
                                Loop
                                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                    ws2.Cells(x1, tbrng + 10).Value = "Test Queue"
                                    x1 = x1 + 1
                                Loop
                                GoTo nn:
                            Else
                                Do While z > 0
                                    ws2.Cells(x - z, tbrng + 6).Value = ""
                                    If x1 = x - z Then ws2.Cells(x1, tbrng + 6).Value = "Test Queue_IQC"
                                    z = z - 1
                                Loop
                                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                    ws2.Cells(x1, tbrng + 6).Value = "Test Queue_IQC"
                                    x1 = x1 + 1
                                Loop
                                GoTo nn:
                            End If
                        End If
                    Loop
                    If Left(ws2.Cells(x, 12), 9) = "Ran Queue" Then
                        If boss = 1 Then
                            ws2.Cells(x, tbrng + 6).Value = "Test Queue_IQC"
                        Else
                            ws2.Cells(x, tbrng + 10).Value = "Test Queue"
                        End If
                    End If
                    If Left(ws2.Cells(x1 - 1, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = "" Then
                        If boss = 1 Then
                            Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                                ws2.Cells(x1 - x2, tbrng + 6).Value = "Test Queue_IQC"
                                x2 = x2 + 1
                            Loop
                        Else
                            Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                                ws2.Cells(x1 - x2, tbrng + 10).Value = "Test Queue"
                                x2 = x2 + 1
                            Loop
                        End If
                    End If
                End If
            End If
            boss = 0
            x = x + 1
            z = 0
            x2 = 1
nn:
        Loop

'Private Sub BQC()
        x = 2
        z = 0
        x1 = 0
        x2 = 1
        'ws2.Range("AB2:AB999999").ClearContents
        ws2.Cells(1, tbrng + 11).Value = "BQC" & Chr(13) & "(Beam Quality Check)"
        Do While x < total
            If Left(ws2.Cells(x, 11), 22) = "OK (linked to IL-0046)" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 11).Value = "BQC"
                    x1 = x1 - 1
                Loop
                Else
                If Left(ws2.Cells(x, 11), 14) = "EVENT: IL-0046" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 22) <> "OK (linked to IL-0046)"
                        ws2.Cells(x, tbrng + 11).Value = "BQC"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 14) = "EVENT: IL-0046" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 11).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 11).Value = "BQC"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 11).Value = "BQC"
                                x1 = x1 + 1
                            Loop
                            GoTo aa:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 22) = "OK (linked to IL-0046)" Then ws2.Cells(x, tbrng + 11).Value = "BQC"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 11).Value = "BQC"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x2 = 1
            x1 = 1
            x = x + 1
            z = 0
aa:
        Loop

'Private Sub DPM_Click()
        x = 2
        z = 0
        x2 = 1
        'ws2.Range("AC2:AC999999").ClearContents
        ws2.Cells(1, tbrng + 12) = "DPM" & Chr(13) & "(Daily Performance Measurement)"
        Do While x < total
            If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: SH-0134" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 12).Value = "DPM"
                    x1 = x1 - 1
                Loop
                Else
                If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: SH-0134" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 27) <> "DEACTIVATE WARNING: SH-0134"
                        ws2.Cells(x, tbrng + 12).Value = "DPM"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: SH-0134" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 12).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 12).Value = "DPM"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 12).Value = "DPM"
                                x1 = x1 + 1
                            Loop
                            GoTo bb:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: SH-0134" Then ws2.Cells(x, tbrng + 12).Value = "DPM"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 12).Value = "DPM"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x = x + 1
            z = 0
            x2 = 1
bb:
        Loop

'Private Sub LCRO_Click()
        i = 2
        x = 2
        x2 = 1
        'ws2.Range("AD2:AD999999").ClearContents
        ws2.Cells(1, tbrng + 13).Value = "LCRO" & Chr(13) & "(Route Calculation)"
        Do While i < total
                For a = 1 To Len(ws2.Cells(x, 12)) + 1
                    If Mid(ws2.Cells(x, 12), a, 4) = "LCRO" Then
                        ws2.Cells(x, tbrng + 13).Value = "LCRO"
                        If Left(ws2.Cells(x + 1, 12), 9) = "Optimised" Then
                            ws2.Cells(x + 1, tbrng + 13).Value = "LCRO"
                                 Do While Left(ws2.Cells(x - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x - x2, 11), 14) = ""
                                    ws2.Cells(x - x2, tbrng + 13).Value = "LCRO"
                                    x2 = x2 + 1
                                Loop
                        End If
                        If Left(ws2.Cells(x - 1, 12), 9) = "Optimised" Then
                            ws2.Cells(x - 1, tbrng + 13).Value = "LCRO"
                                Do While Left(ws2.Cells(x - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x - x2, 11), 14) = ""
                                    ws2.Cells(x - x2 - 1, tbrng + 13).Value = "LCRO"
                                    x2 = x2 + 1
                                Loop
                        End If
                        GoTo kk:
                    End If
                Next a
kk:
        i = i + 1
        x = x + 1
        x2 = 1
        Loop

'Private Sub LCD_Click()
        x = 2
        z = 0
        'ws2.Range("AE2:AE999999").ClearContents
        ws2.Cells(1, tbrng + 14).Value = "LCD" & Chr(13) & "(Determine Exposure/Scandirection)"
        Do While x < total
            For z = 1 To Len(ws2.Cells(x, 11))
            If Mid(ws2.Cells(x, 11), z, 7) = "LC-0015" Then ws2.Cells(x, tbrng + 14).Value = "LCD"
            Next z
            x = x + 1
        Loop
'MsgBox "lazer new fill"
'ALL LASER
        'Private Sub Laser_New_Fill_Click()
        x = 2
        z = 0
        x1 = 0
        x2 = 1
        'ws2.Range("AF2:AI999999").ClearContents
        ws2.Cells(1, tbrng + 15).Value = "Laser: New Fill"
        Do While x < total
            If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: IL-0065" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 15).Value = "Laser: New Fill"
                    x1 = x1 - 1
                Loop
                Else
                If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: IL-0065" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 27) <> "DEACTIVATE WARNING: IL-0065"
                        ws2.Cells(x, tbrng + 15).Value = "Laser: New Fill"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: IL-0065" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 15).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 15).Value = "Laser: New Fill"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 15).Value = "Laser: New Fill"
                                x1 = x1 + 1
                            Loop
                            GoTo hh:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: IL-0065" Then ws2.Cells(x, tbrng + 15).Value = "Laser: New Fill"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 15).Value = "Laser: New Fill"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x = x + 1
            z = 0
            x2 = 1
hh:
        Loop

        'Private Sub Laser_Gas_Mixture_Click()
        x = 2
        z = 0
        x1 = 0
        x2 = 1
        ws2.Cells(1, tbrng + 16).Value = "Laser: Gas Mixture"
        Do While x < total
            If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: IL-00b6" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 16).Value = "Laser: Gas Mixture"
                    x1 = x1 - 1
                Loop
        
                Else
                If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: IL-00b6" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 27) <> "DEACTIVATE WARNING: IL-00b6"
                        ws2.Cells(x, tbrng + 16).Value = "Laser: Gas Mixture"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: IL-00b6" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 16).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 16).Value = "Laser:Gas Mixture"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 16).Value = "Laser:Gas Mixture"
                                x1 = x1 + 1
                            Loop
                            GoTo ff:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: IL-00b6" Then ws2.Cells(x, tbrng + 16).Value = "Laser: Gas Mixture"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 16).Value = "Laser: Gas Mixture"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x2 = 1
            x = x + 1
            z = 0
ff:
        Loop

        'Private Sub Laser_Halogen_Click()
        x = 2
        z = 0
        x1 = 0
        x2 = 1
        ws2.Cells(1, tbrng + 17).Value = "Laser: Halogen"
        Do While x < total
            If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: IL-00ae" Then
                x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 17).Value = "Laser: Halogen"
                    x1 = x1 - 1
                Loop
                Else
                If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: IL-00ae" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 27) <> "DEACTIVATE WARNING: IL-00ae"
                        ws2.Cells(x, tbrng + 17).Value = "Laser: Halogen"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 25) = "ACTIVATE WARNING: IL-00ae" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 17).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 17).Value = "Laser: Halogen"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 17).Value = "Laser: Halogen"
                                x1 = x1 + 1
                            Loop
                            GoTo gg:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 27) = "DEACTIVATE WARNING: IL-00ae" Then ws2.Cells(x, tbrng + 17).Value = "Laser: Halogen"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 17).Value = "Laser: Halogen"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x2 = 1
            x = x + 1
            z = 0
gg:
        Loop
        
        'Private Sub Laser_Others_Click()
        x = 2
        z = 0
        ws2.Cells(1, tbrng + 18).Value = "Laser: Others"
        Do While x < total
            For z = 1 To Len(ws2.Cells(x, 12))
                If Mid(ws2.Cells(x, 12), z, 5) = "Laser" Then
                    If ws2.Cells(x, tbrng + 15) = "" And ws2.Cells(x, tbrng + 16) = "" And ws2.Cells(x, tbrng + 17) = "" Then
                        ws2.Cells(x, tbrng + 18).Value = "Laser: Others"
                        GoTo ii:
                    End If
                End If
            Next z
ii:
        x = x + 1
        Loop

'Private Sub LightSource_Click()
        x = 2
        z = 0
        x1 = 0
        x2 = 1
        'ws2.Range("AJ2:AJ999999").ClearContents
        ws2.Cells(1, tbrng + 19).Value = "Light Source"
        Do While x < total
            If Left(ws2.Cells(x, 11), 14) = "EVENT: IL-0118" Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 19).Value = "Lightsource"
                    x1 = x1 - 1
                Loop
                Else
                If Left(ws2.Cells(x, 11), 14) = "EVENT: IL-0119" Then
                x1 = x
                    Do While Left(ws2.Cells(x, 11), 14) <> "EVENT: IL-0118"
                        ws2.Cells(x, tbrng + 19).Value = "Lightsource"
                        x = x + 1
                        z = z + 1
                        If Left(ws2.Cells(x, 11), 14) = "EVENT: IL-0119" Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 19).Value = ""
                                If x1 = x - z Then ws2.Cells(x1, tbrng + 19).Value = "Lightsource"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 19).Value = "Lightsource"
                                x1 = x1 + 1
                            Loop
                            GoTo ll:
                        End If
                    Loop
                    If Left(ws2.Cells(x, 11), 14) = "EVENT: IL-0118" Then ws2.Cells(x, tbrng + 19).Value = "Lightsource"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        ws2.Cells(x1 - x2, tbrng + 19).Value = "Lightsource"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x2 = 1
            x = x + 1
            z = 0
ll:
        Loop

'MsgBox "a"

'Private Sub Pause_Click()
        x = 2
        z = 0
        x2 = 1
        'ws2.Range("AC2:AC999999").ClearContents
        ws2.Cells(1, tbrng + 20) = "Pause"
        Do While x < total
            If Len(ws2.Cells(x, 12)) <> Len(Replace(ws2.Cells(x, 12), "Online / Remote", "")) Then
            x1 = x
                Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                    ws2.Cells(x1, tbrng + 20).Value = "Pause"
                    x1 = x1 - 1
                Loop
                Else
                If Len(ws2.Cells(x, 12)) <> Len(Replace(ws2.Cells(x, 12), "Online / Local", "")) Then
                x1 = x
                'MsgBox Len(ws2.Cells(x, 12)) & " " & Len(Replace(ws2.Cells(x, 12), "Online / Remote", ""))
                'If 36 <> 36 Then MsgBox "a"
                    Do Until CInt(Len(ws2.Cells(x, 12))) <> CInt(Len(Replace(ws2.Cells(x, 12), "Online / Remote", "")))
                        If Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss" Then ws2.Cells(x, tbrng + 20).Value = "Pause"
                        x = x + 1
                        z = z + 1
                        If Len(ws2.Cells(x, 12)) <> Len(Replace(ws2.Cells(x, 12), "Online / Local", "")) Or x = total Then
                            Do While z > 0
                                ws2.Cells(x - z, tbrng + 20).Value = ""
                                If x1 = x - z And Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss" Then ws2.Cells(x1, tbrng + 20).Value = "Pause"
                                z = z - 1
                            Loop
                            Do While Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss"
                                ws2.Cells(x1, tbrng + 20).Value = "Pause"
                                x1 = x1 + 1
                            Loop
                            GoTo oo:
                        End If
                    Loop
                    If Len(ws2.Cells(x, 12)) <> Len(Replace(ws2.Cells(x, 12), "Online / Remote", "")) And Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss" Then ws2.Cells(x, tbrng + 20).Value = "Pause"
                    Do While Left(ws2.Cells(x1 - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x1 - x2, 11), 14) = ""
                        If Left(ws2.Cells(x1, tbrng + 1), 4) = "Loss" Then ws2.Cells(x1 - x2, tbrng + 20).Value = "Pause"
                        x2 = x2 + 1
                    Loop
                End If
            End If
            x = x + 1
            z = 0
            x2 = 1
oo:
        Loop

'dbf:

'Private Sub Priority_Click()
        boss = 0
finalstate:
        'MsgBox boss
        x = 2
        x1 = 0
        ws2.Cells(1, tbrng + 21).Value = "Priority"
        ws2.Range(ws2.Cells(2, tbrng + 21), ws2.Cells(999999, tbrng + 21)).ClearContents
        'If boss = 0 Then ws2.Range(ws2.Cells(2, tbrng + 4), ws2.Cells(999999, tbrng + 4)).ClearContents
        Do While x < total
            'If x = 180 Then MsgBox x
            y = tbrng + 2 'max 32, target 34
            Do While y <= tbrng + 20
                If ws2.Cells(x, y) <> "" Then
                    text = ws2.Cells(x, y)
                    For z = 1 To Len(text)
                        If Mid(text, z, 1) <> " " Then text1 = text1 + Mid(text, z, 1)
                    Next z
                    ws2.Cells(x, tbrng + 21).Value = text1
                    
                    If boss = 1 Then
                        If Left(text1, 13) = "IntrinsicLoss" And CStr(text1) <> "IntrinsicLoss_Sxxx" Then
                            x1 = x
                        Else
                            x1 = 0
                        End If
                    End If
                    GoTo mm:
                Else
                    y = y + 1
                End If
            Loop
            If boss = 1 And x1 <> 0 And ws2.Cells(x, tbrng + 1) = "Loss" Then ws2.Cells(x, tbrng + 21).Value = ws2.Cells(x1, tbrng + 21)
mm:
            x = x + 1
            text = ""
            text1 = ""
        Loop
        ws2.Range(ws2.Cells(2, tbrng + 21), ws2.Cells(999999, tbrng + 21)).Interior.ColorIndex = 0
        ws2.Range(ws2.Cells(2, tbrng + 21), ws2.Cells(999999, tbrng + 21)).Font.Color = RGB(0, 0, 0)
        If boss = 1 Then GoTo endalr:
        
'Private Sub Logic3_Click()
        x = 2
        x1 = 0
        boss = 0
        x2 = 1
        ws2.Cells(1, tbrng + 5).Value = "Logic3" & Chr(13) & "(Instrinsic Loss)"
        Do While x < total
            If ws2.Cells(x, tbrng + 21) <> "" Then boss = 0
            If Left(ws2.Cells(x, 11), 14) = "EVENT: BC-0050" And ws2.Cells(x, tbrng + 1) = "Loss" Then
                Select Case Left(ws2.Cells(x, 8), 1)
                    Case "S"
                        Select Case Mid(ws2.Cells(x, 8), 2, 1)
                            Case "5"
                                ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss_S5"
                                    x1 = x
                                    boss = 1
                            Case "6"
                                ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss_S6"
                                    x1 = x
                                    boss = 1
                            Case "9"
                                ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss_S9"
                                    x1 = x
                                    boss = 1
                            Case Else
                                If boss = 1 Then
                                    If Left(ws2.Cells(x1, 8), 1) = "S" Or Left(ws2.Cells(x1, 8), 1) = "s" Then
                                    ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss" + "_" + Left(ws2.Cells(x1, 8), 2)
                                    Else
                                    ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss" + "_" + Left(ws2.Cells(x1, 8), 1)
                                    End If
                                    boss = 0
                                Else
                                    ws2.Cells(x, tbrng + 5).Value = _
                                        "Intrinsic Loss_Sxxx"
                                End If
                        End Select
                        
                    Case "s"
                        Select Case Mid(ws2.Cells(x, 8), 2, 1)
                            Case "5"
                                ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss_S5"
                                    x1 = x
                                    boss = 1
                            Case "6"
                                ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss_S6"
                                    x1 = x
                                    boss = 1
                            Case "9"
                                ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss_S9"
                                    x1 = x
                                    boss = 1
                            Case Else
                                If boss = 1 Then
                                    If Left(ws2.Cells(x1, 8), 1) = "S" Or Left(ws2.Cells(x1, 8), 1) = "s" Then
                                    ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss" + "_" + Left(ws2.Cells(x1, 8), 2)
                                    Else
                                    ws2.Cells(x, tbrng + 5).Value = _
                                    "Intrinsic Loss" + "_" + Left(ws2.Cells(x1, 8), 1)
                                    End If
                                    boss = 0
                                Else
                                    ws2.Cells(x, tbrng + 5).Value = _
                                        "Intrinsic Loss_Sxxx"
                                End If
                        End Select
                        
                    Case "Z"
                        ws2.Cells(x, tbrng + 5).Value = _
                            "Intrinsic Loss_Z"
                            x1 = x
                            boss = 1
                    Case "z"
                        ws2.Cells(x, tbrng + 5).Value = _
                            "Intrinsic Loss_Z"
                            x1 = x
                            boss = 1
                    Case "C"
                        ws2.Cells(x, tbrng + 5).Value = _
                            "Intrinsic Loss" + "_" + Left(ws2.Cells(x, 8), 1)
                            x1 = x
                            boss = 1
                    Case "M"
                        ws2.Cells(x, tbrng + 5).Value = _
                            "Intrinsic Loss" + "_" + Left(ws2.Cells(x, 8), 1)
                            x1 = x
                            boss = 1
                    Case Else
                        ws2.Cells(x, tbrng + 5).Value = _
                            "Intrinsic Loss_Others"
                End Select
                Do While Left(ws2.Cells(x - x2, 11), 14) = "EVENT: BC-0059" Or Left(ws2.Cells(x - x2, 11), 14) = ""
                    ws2.Cells(x - x2, tbrng + 5).Value = ws2.Cells(x, tbrng + 5)
                    x2 = x2 + 1
                Loop
            End If
        x = x + 1
        x2 = 1
        Loop

boss = 1
GoTo finalstate:
endalr:


db:
'Private Sub Temp&Resist&Position_Click()
'tbrng = 19

        ws2.Cells(1, tbrng + 22).Value = "Temperature"
        ws2.Cells(1, tbrng + 23).Value = "Resist"
        ws2.Cells(1, tbrng + 24).Value = "Position"
        ws2.Cells(1, tbrng + 28).Value = "Barc"
        For i = 2 To total
            text = ws2.Cells(i, tbrng - 1)
            If Len(text) - Len(Replace(text, "-", "")) <> 0 Then
            a = 1
            text = ""
            text1 = ""
            On Error Resume Next
                Do Until a > 2 'taking resist
                    If CInt(Mid(ws2.Cells(i, tbrng - 1), InStr(ws2.Cells(i, tbrng - 1), "-") - a, 1)) < 0 Then
                    text1 = Mid(ws2.Cells(i, tbrng - 1), InStr(ws2.Cells(i, tbrng - 1), "-") - a, 1) & text1
                    End If
                    a = a + 1
                Loop
                a = 1
                text = Mid(ws2.Cells(i, tbrng - 1), InStr(ws2.Cells(i, tbrng - 1), "-") + 3, 1)
                    a = a + 1
                If CInt(Mid(ws2.Cells(i, tbrng - 1), InStr(ws2.Cells(i, tbrng - 1), "-") + 2 + a, 1)) < 0 Then 'taking temp
                    text = text & Mid(ws2.Cells(i, tbrng - 1), InStr(ws2.Cells(i, tbrng - 1), "-") + 2 + a, 1)
                End If
                
                ws2.Cells(i, tbrng + 22).Value = WorksheetFunction.VLookup(CInt(text), wst.Range(wst.Columns(1), wst.Columns(2)), 2, 0) '"Temperature"
                ws2.Cells(i, tbrng + 23).Value = WorksheetFunction.VLookup(text1, wsr.Range(wsr.Columns(1), wsr.Columns(2)), 2, 0) '"Resist"
                ws2.Cells(i, tbrng + 24).Value = WorksheetFunction.VLookup(ws2.Cells(i, tbrng + 22) & Left(ws2.Cells(i, 4), 6), wsp.Range(wsp.Columns(3), wsp.Columns(4)), 2, 0) '"Position"
'=IF(AG2="W","SP25 ME",IF(LEFT(AF2,1)="0","Normal",IF(LEFT(AF2,1)="1","OL Barc",IF(LEFT(AF2,1)="G","IL Barc",""))))
                'MsgBox "a"
                Select Case CStr(Mid(ws2.Cells(i, tbrng - 1), 5, 1))
                    Case "W"
                        ws2.Cells(i, tbrng + 28).Value = "SP25 ME"
                    Case Else
                        Select Case CStr(Left(ws2.Cells(i, tbrng - 1), 1))
                            Case "0"
                                ws2.Cells(i, tbrng + 28).Value = "Normal"
                            Case "1"
                                ws2.Cells(i, tbrng + 28).Value = "OL Barc"
                            Case "G"
                                ws2.Cells(i, tbrng + 28).Value = "IL Barc"
                        End Select
                End Select
            End If
        Next i


'Private Sub Tracking_Click()


        ws2.Cells(1, tbrng + 25).Value = "Temp_Change"
        ws2.Cells(1, tbrng + 26).Value = "Res_Change"
        ws2.Cells(1, tbrng + 27).Value = "Pos_Change"
        ws2.Cells(1, tbrng + 29).Value = "Classification"
        'tbrng = 19
        x = 0
        x1 = 0
        x2 = 0
        For i = 2 To total
            If ws2.Cells(i, tbrng + 22) <> "" Then
                If x = 0 Then
                    x = i
                Else
                    Select Case Left(ws2.Cells(i, tbrng + 22) - ws2.Cells(x, tbrng + 22), 1)
                        Case "-"
                            ws2.Cells(i, tbrng + 25).Value = "Cooler (" & Right(ws2.Cells(i, tbrng + 22) - ws2.Cells(x, tbrng + 22), Len(ws2.Cells(i, tbrng + 22) - ws2.Cells(x, tbrng + 22)) - 1) & ")"
                        Case "0"
                            ws2.Cells(i, tbrng + 25).Value = "No Change"
                        Case Else
                            ws2.Cells(i, tbrng + 25).Value = "Warmer (" & ws2.Cells(i, tbrng + 22) - ws2.Cells(x, tbrng + 22) & ")"
                    End Select
                    x = i
                End If
            End If
            If ws2.Cells(i, tbrng + 23) <> "" Then
                If x1 = 0 Then
                    x1 = i
                Else
                    If ws2.Cells(i, tbrng + 23) <> ws2.Cells(x1, tbrng + 23) Then
                        ws2.Cells(i, tbrng + 26).Value = ws2.Cells(x1, tbrng + 23) & " To " & ws2.Cells(i, tbrng + 23)
                    Else
                        ws2.Cells(i, tbrng + 26).Value = "No Change"
                    End If
                    x1 = i
                End If
            End If
            If ws2.Cells(i, tbrng + 24) <> "" Then
                If x2 = 0 Then
                    x2 = i
                Else
                    If ws2.Cells(i, tbrng + 24) <> ws2.Cells(x2, tbrng + 24) Then
                        ws2.Cells(i, tbrng + 27).Value = ws2.Cells(x2, tbrng + 24) & " To " & ws2.Cells(i, tbrng + 24)
                    Else
                        ws2.Cells(i, tbrng + 27).Value = "No Change"
                    End If
                    x2 = i
                End If
            End If
            If ws2.Cells(i, tbrng + 28) <> "" Then
                If x3 = 0 Then
                    x3 = i
                Else
                    If ws2.Cells(i, tbrng + 28) <> ws2.Cells(x3, tbrng + 28) Then
                        ws2.Cells(i, tbrng + 29).Value = ws2.Cells(x3, tbrng + 28) & " To " & ws2.Cells(i, tbrng + 28)
                    Else
                        ws2.Cells(i, tbrng + 29).Value = "No Change"
                    End If
                    x3 = i
                End If
            End If

        Next i


ws2.Range(ws2.Cells(1, tbrng + 3), ws2.Cells(1, tbrng + 29)).Font.Bold = True
ws2.Range(ws2.Cells(1, tbrng + 3), ws2.Cells(1, tbrng + 29)).Interior.ColorIndex = 6
ws2.Range(ws2.Cells(1, tbrng + 1), ws2.Cells(WorksheetFunction.CountA(ws2.Columns(tbrng + 1)), tbrng + 1)).Interior.ColorIndex = 45



ws2.Cells(total, tbrng + 1).Value = "Loss"
ws2.Cells(total, tbrng + 3).Value = "TEST_Loss"
ws2.Cells(total, tbrng + 4).Value = "Cold Start"""
ws2.Cells(total, tbrng + 5).Value = "Intrinsic Loss_Sxxx"
ws2.Cells(total, tbrng + 6).Value = "Test Queue_IQC"
ws2.Cells(total, tbrng + 7).Value = "IRIS"
ws2.Cells(total, tbrng + 8).Value = "No Reticle"
ws2.Cells(total, tbrng + 9).Value = "IQC"
ws2.Cells(total, tbrng + 10).Value = "Test Queue"
ws2.Cells(total, tbrng + 11).Value = "BQC"
ws2.Cells(total, tbrng + 12).Value = "DPM"
ws2.Cells(total, tbrng + 13).Value = "LCRO"
ws2.Cells(total, tbrng + 14).Value = "LCD"
ws2.Cells(total, tbrng + 15).Value = "Laser: New Fill"
ws2.Cells(total, tbrng + 16).Value = "Laser: Gas Mixture"
ws2.Cells(total, tbrng + 17).Value = "Laser: Halogen"
ws2.Cells(total, tbrng + 18).Value = "Laser: Others"
ws2.Cells(total, tbrng + 19).Value = "Lightsource"
ws2.Cells(total, tbrng + 20).Value = "Pause"
ws2.Cells(total, tbrng + 21).Value = "TEST_Loss"

ThisWorkbook.Save

'MsgBox "Done!", , "Project_Analyze"
End Sub
