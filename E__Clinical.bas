Attribute VB_Name = "E__Clinical"
'WATSON COLORS
'----------------------
'Color = RGB(180, 0, 0)      '<<OUT OF SCOPE>>
'color = 65280              '<<IN SCOPE>>
'color = 652804             '<<TBD>>
'********************************************************
'SHERLOCK COLORS
'----------------------
'Color = RGB(254, 0, 0)      '<<OUT OF SCOPE>>
'Color = RGB(0, 127, 0)      '<<IN SCOPE>>
'Color = RGB(254, 254, 0)    '<<TBD>>
'********************************************************


Sub StrtSherlock(control As IRibbonControl)

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
    
SherlockForm.Show False

End Sub
Sub clinicalrun()

Application.ScreenUpdating = False

    If Not Range("ZK1").Value = "1" Then        'skip for viewall
        Call parseKeywords  '>>>>>>>>>>
        Call PSCxxOption    '>>>>>>>>>>
        Call PFfind         '>>>>>>>>>>
    End If

    'Count total rows
    '-------------------------------------------
    Sheets("Line Item Data").Select
    Range("H3").Select
    'lastRow = Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    Range(Selection, Selection.End(xlDown)).Select
    ttlrows = Application.WorksheetFunction.CountA(Selection) - 1
    
    'Dim Sh As Worksheet, flg As Boolean
    For Each sht In Worksheets
        If sht.Name = "ClinicalQC" Then
            Range("ZA1").Value = 1
            Columns("A:N").Hidden = False
            Call FUN_TestForSheet("ClinicalQC")
            GoTo QCShtPresent
        End If
    Next
    
    Range("ZA1").Value = 1
    Columns("A:M").Hidden = False
    Sheets.Add After:=Sheets("Line Item Data")
    ActiveSheet.Name = "ClinicalQC"
    Range("A1").Value = "Original Sort"
    Range("B1").Value = "PSC"
    Range("C1").Value = "PIM key"
    Range("D1").Value = "Description"
    Range("E1").Value = "Manufacturer"
    Range("F1").Value = "Standard Manufacturer"
    Range("G1").Value = "Catalog Number"
    Range("H:H").Interior.ColorIndex = 1
    Range("H:H").ColumnWidth = 3
    Range("I1").Value = "Pricefile catalog number"
    Range("J1").Value = "Pricefile Description"
    Range("K1").Value = "Supplier Name"
    Range("A1:R1").Interior.ColorIndex = 15
    Range("L:L").Interior.ColorIndex = 1
    Range("L:L").ColumnWidth = 3

QCShtPresent:
    'reset clinical values
    '-------------------------------------------
    Range("A2:G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Sheets("Line Item Data").Select
    
    'return results (filter)
    '-------------------------------------------
    If Not Application.WorksheetFunction.sum(Range("ZL:ZL")) = 0 Then
    On Error GoTo errhndlefilter
    ActiveSheet.Range(Range("A2"), Range("ZL2").Offset(ttlrows, 0)).AutoFilter Field:=688, Criteria1:="<>"  'Filter on (ZL)
    On Error GoTo 1
    Range(Range("A3"), Range("A3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    On Error GoTo 0
    Sheets("clinicalqc").Select
    Range("a2").Select
    ActiveSheet.Paste
    Sheets("Line Item Data").Select
    Range(Range("C3"), Range("C3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("clinicalqc").Select
    Range("B2").Select
    ActiveSheet.Paste
    Sheets("Line Item Data").Select
    Range(Range("E3"), Range("E3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("clinicalqc").Select
    Range("C2").Select
    ActiveSheet.Paste
    Sheets("Line Item Data").Select
    Range(Range("H3"), Range("H3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("clinicalqc").Select
    Range("D2").Select
    ActiveSheet.Paste
    Sheets("Line Item Data").Select
    Range(Range("K3"), Range("K3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("clinicalqc").Select
    Range("E2").Select
    ActiveSheet.Paste
    Sheets("Line Item Data").Select
    Range(Range("L3"), Range("L3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("clinicalqc").Select
    Range("F2").Select
    ActiveSheet.Paste
    Sheets("Line Item Data").Select
    Range(Range("N3"), Range("N3").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("clinicalqc").Select
    Range("G2").Select
    ActiveSheet.Paste
    Columns("A:K").EntireColumn.AutoFit
    If Sheets("Line Item Data").Range("ZK1").Value = 1 Then
        Range("M1:R1").Interior.ColorIndex = 0
        Call SherlockRunKeys  '>>>>>>>>>>
        On Error GoTo 1
        Range("M:M").SpecialCells(xlCellTypeConstants, 1).Select
        On Error GoTo 0
        Selection.Delete Shift:=xlUp
        Range("M1:R1").Insert Shift:=xlDown
        Range("M1").Value = "Single Word"
        Range("O1").Value = "Two Word"
        Range("Q1").Value = "Three Word"
        Range("M1:R1").Interior.ColorIndex = 15
        Sheets("Line Item Data").Range("ZK1").ClearContents
    End If
    
    On Error Resume Next
    Sheets("Line Item Data").ShowAllData
    Exit Sub
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::;;
    End If
1   Sheets("clinicalqc").Select
    Columns("A:K").EntireColumn.AutoFit
    Sheets("Line Item Data").Range("ZK1").ClearContents
    'If Not FUN_lastrow("A") > 2000 Then
    On Error Resume Next
    Sheets("Line Item Data").ShowAllData
    'End If
    Exit Sub

errhndlefilter:
    Rows("2:2").AutoFilter
    On Error GoTo 0
    Resume
    



End Sub
Sub parseKeywords()

Application.ScreenUpdating = False

critnmbr = 0
Range("ZM:ZZ").ClearContents

'Parse PSC
'=======================================================================
parskey = SherlockForm.PSCtxtform.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "PSC"    '(placeholder for selecting down)

    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "PSC"    '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If
critnmbr = critnmbr + 1

'Parse PIM
'=======================================================================
parskey = SherlockForm.PIMtxtform.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "PIM"    '(placeholder for selecting down)
    
    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "PIM"    '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If
critnmbr = critnmbr + 1

'Parse Desc
'=======================================================================
parskey = SherlockForm.DESCtxtform.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "Desc"    '(placeholder for selecting down)

    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "Desc"    '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If
critnmbr = critnmbr + 1

'Parse Mfg
'=======================================================================
parskey = SherlockForm.MFGtxtform.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "Mfg"    '(placeholder for selecting down)

    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "Mfg"    '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If
critnmbr = critnmbr + 1

'Parse Catnum
'=======================================================================
parskey = SherlockForm.CATNUMtxtform.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "Catnum"    '(placeholder for selecting down)

    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "Catnum"    '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If
critnmbr = critnmbr + 1

'Parse PFkeyword
'=======================================================================
parskey = SherlockForm.PFkeyword.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "PFkeyword"    '(placeholder for selecting down)

    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "PFkeyword"    '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If
critnmbr = critnmbr + 1

'Parse PFcatnum
'=======================================================================
parskey = SherlockForm.PFcatnum.Text

If InStr(parskey, ";") > 0 Then
    Range("ZM2").Offset(0, critnmbr).Value = "PFcatnum"     '(placeholder for selecting down)
    
    'count number of ;
    '----------------------
    chrcount = 0
    For c = 1 To Len(parskey)
        If Mid(parskey, c, 1) = ";" Then
            chrcount = chrcount + 1
        End If
    Next
    
    'Parse Words
    '----------------------
    parscount = 0
    Do
        parscount = parscount + 1
        On Error GoTo errhndlParskey
        Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
        parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
        On Error GoTo 0
    Loop Until parscount = chrcount + 1
Else
    Range("ZM2").Offset(0, critnmbr).Value = "PFcatnum"     '(placeholder for selecting down)
    Sheets("Line Item Data").Range("ZM2").Offset(1, critnmbr).Value = parskey
End If



Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlParskey:
Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = Trim(parskey)
Resume Next

'errhndlnmbrParskey:
'Sheets("Line Item Data").Range("ZM2").Offset(parscount, critnmbr).Value = "A" & Trim(parskey)
'Resume Next


End Sub
Sub PSCxxOption()

Application.ScreenUpdating = False

'(Does it match) PSC Contains / does not contain (returns if matches any keywords)
'-------------------------------
    If SherlockForm.PSCtxtform.Text = "" And SherlockForm.PSCand <> True And SherlockForm.PSCor <> True Then
    Else
        If SherlockForm.PSCcheck.Value = False Then 'Or PSCcheck = Null Then
            Range("ZA3").Formula = "=if(countif(" & Range(Range("ZM3"), Range("ZM2").End(xlDown)).Address & "," & Range("C3").Address(0, 0) & ")=1,1,"""")"      '(mark with 1)
            Range("ZA3").AutoFill Destination:=Range("ZA3:ZA" & Cells(Rows.Count, "N").End(xlUp).Row)
            Range("ZA3:ZA" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
        Else
            Range("ZA3").Formula = "=if(countif(" & Range(Range("ZM3"), Range("ZM2").End(xlDown)).Address & "," & Range("C3").Address(0, 0) & ")<>1,1,"""")"       '(mark with 1)
            Range("ZA3").AutoFill Destination:=Range("ZA3:ZA" & Cells(Rows.Count, "N").End(xlUp).Row)
            Range("ZA3:ZA" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
        End If
    
        '(Should it be included) PSC and / or
        '-------------------------------
        If SherlockForm.PSCand = True Then
            Range("ZB3").Formula = "=if(ZA3=1," & Chr(34) & "y" & Chr(34) & "," & Chr(34) & "x" & Chr(34) & ")"                                            '(mark with 1)
            Range("ZB3").AutoFill Destination:=Range("ZB3:ZB" & Cells(Rows.Count, "N").End(xlUp).Row)
            Range("ZB3:ZB" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
        ElseIf SherlockForm.PSCdna <> True Then
            Range("ZB3").Formula = "=if(ZA3=1," & Chr(34) & "y" & Chr(34) & ","""")"                                                            '(mark with 1)
            Range("ZB3").AutoFill Destination:=Range("ZB3:ZB" & Cells(Rows.Count, "N").End(xlUp).Row)
            Range("ZB3:ZB" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
        End If
    End If


'(Does it match) PIM Contains / does not contain (returns if matches any PIM)
'-------------------------------
If SherlockForm.PIMtxtform.Text = "" And SherlockForm.PIMand <> True And SherlockForm.PIMor <> True Then
Else
    If SherlockForm.PIMcheck.Value = False Then
        Range("ZC3").Formula = "=if(countif(" & Range(Range("ZN3"), Range("ZN2").End(xlDown)).Address & "," & Range("E3").Address(0, 0) & ")=1,1,"""")"         '(mark with 1)
        Range("ZC3").AutoFill Destination:=Range("ZC3:ZC" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZC3:ZC" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    Else
        Range("ZC3").Formula = "=if(countif(" & Range(Range("ZN3"), Range("ZN2").End(xlDown)).Address & "," & Range("E3").Address(0, 0) & ")<>1,1,"""")"         '(mark with 1)
        Range("ZC3").AutoFill Destination:=Range("ZC3:ZC" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZC3:ZC" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If
    
    '(Should it be included) PIM and / or
    '-------------------------------
    If SherlockForm.PIMand = True Then
        Range("ZD3").Formula = "=if(ZA3=1," & Chr(34) & "y" & Chr(34) & "," & Chr(34) & "x" & Chr(34) & ")"                                                             '(mark with 1)
        Range("ZD3").AutoFill Destination:=Range("ZD3:ZD" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZD3:ZD" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    ElseIf SherlockForm.PIMdna <> True Then
        Range("ZD3").Formula = "=if(ZA3=1," & Chr(34) & "y" & Chr(34) & ","""")"                                                            '(mark with 1)
        Range("ZD3").AutoFill Destination:=Range("ZD3:ZD" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZD3:ZD" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If
End If

'(Does it match) Desc Contains / does not contain  (returns if matches all keywords)
'-------------------------------
If SherlockForm.DESCtxtform.Text = "" And SherlockForm.DESCand <> True And SherlockForm.DESCor <> True Then
Else
    If SherlockForm.DescCheck.Value = False Then    '(is an "and" search, must find both words to be valid)
        DescCnt = 0
        Do
            DescCnt = DescCnt + 1
            DoEvents
            Application.StatusBar = "Checking HCO keywords: row " & (Range("H2").Offset(DescCnt, 0).Row)
            innercnt = 0
            'mtchcnt = 0
            For Each c In Range(Range("ZO3"), Range("ZO2").End(xlDown))
                innercnt = innercnt + 1
                'Debug.Print LCase(Range("H2").Offset(desccnt, 0).Value)
                'Debug.Print LCase(Range("ZO2").Offset(innercnt, 0).Value)
                If InStr(LCase(Range("H2").Offset(DescCnt, 0).Value), LCase(Range("ZO2").Offset(innercnt, 0).Value)) = 0 Then
                    'mtchcnt = mtchcnt + 1
                    GoTo 5
                End If
            Next
            'If mtchcnt = Application.CountA(Range(Range("ZO3"), Range("ZO2").End(xlDown))) Then
                'Range("ZE2").Offset(desccnt, 0).Value = 1
            'End If
            Range("ZE2").Offset(DescCnt, 0).Value = 1
5       Loop Until IsEmpty(Range("A2").Offset(DescCnt, 0))
    Else     '(is an "or" search, if finds either of the words listed it kicks it out)
        DescCnt = 0
        Do
            DescCnt = DescCnt + 1
            Application.StatusBar = "Checking HCO keywords: row " & Range("A1").End(xlDown).Row - (Range("H2").Offset(DescCnt, 0).Row)
            innercnt = 0
            'mtchcnt = 0
            For Each c In Range(Range("ZO3"), Range("ZO2").End(xlDown))
                innercnt = innercnt + 1
                'Debug.Print LCase(Range("H2").Offset(desccnt, 0).Value)
                'Debug.Print LCase(Range("ZO2").Offset(innercnt, 0).Value)
                If InStr(LCase(Range("H2").Offset(DescCnt, 0).Value), LCase(Range("ZO2").Offset(innercnt, 0).Value)) > 0 Then
                    
                    GoTo 6
                    
                End If
            Next
            'If mtchcnt = Application.CountA(Range(Range("ZO3"), Range("ZO2").End(xlDown))) Then
                Range("ZE2").Offset(DescCnt, 0).Value = 1
            'End If
6       Loop Until IsEmpty(Range("A2").Offset(DescCnt, 0))
    End If
    
    '(Should it be included) Desc and / or
    '-------------------------------
    If SherlockForm.DESCand = True Then
        Range("ZF3").Formula = "=if(ZE3=1," & Chr(34) & "y" & Chr(34) & "," & Chr(34) & "x" & Chr(34) & ")"                                                             '(mark with 1)
        Range("ZF3").AutoFill Destination:=Range("ZF3:ZF" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZF3:ZF" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    ElseIf SherlockForm.DESCdna <> True Then
        Range("ZF3").Formula = "=if(ZE3=1," & Chr(34) & "y" & Chr(34) & ","""")"                                                            '(mark with 1)
        Range("ZF3").AutoFill Destination:=Range("ZF3:ZF" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZF3:ZF" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If
End If

'(Does it match) Mfg Contains / does not contain
'-------------------------------
If SherlockForm.MFGtxtform.Text = "" And SherlockForm.MFGand <> True And SherlockForm.MFGor <> True Then
Else
    If SherlockForm.MfgCheck.Value = False Then
        Range("ZG3").Formula = "=if(countif(" & Range(Range("ZP3"), Range("ZP2").End(xlDown)).Address & "," & Range("L3").Address(0, 0) & ")=1,1,"""")"         '(mark with 1)
        Range("ZG3").AutoFill Destination:=Range("ZG3:ZG" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZG3:ZG" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    Else
        Range("ZG3").Formula = "=if(countif(" & Range(Range("ZP3"), Range("ZP2").End(xlDown)).Address & "," & Range("L3").Address(0, 0) & ")<>1,1,"""")"         '(mark with 1)
        Range("ZG3").AutoFill Destination:=Range("ZG3:ZG" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZG3:ZG" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If
    
    '(Should it be included) Mfg and / or
    '-------------------------------
    If SherlockForm.MFGand = True Then
        Range("ZH3").Formula = "=if(ZG3=1," & Chr(34) & "y" & Chr(34) & "," & Chr(34) & "x" & Chr(34) & ")"                                                             '(mark with 1)
        Range("ZH3").AutoFill Destination:=Range("ZH3:ZH" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZH3:ZH" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    ElseIf SherlockForm.MFGdna <> True Then
        Range("ZH3").Formula = "=if(ZG3=1," & Chr(34) & "y" & Chr(34) & ","""")"                                                           '(mark with 1)
        Range("ZH3").AutoFill Destination:=Range("ZH3:ZH" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZH3:ZH" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If
End If

'(Does it match) Catnum Contains / does not contain
'-------------------------------
If SherlockForm.CATNUMtxtform.Text = "" And SherlockForm.CATNUMand <> True And SherlockForm.CATNUMor <> True Then
Else
    If SherlockForm.CatnumCheck.Value = False Then
        Range("ZI3").Formula = "=if(countif(" & Range(Range("ZQ3"), Range("ZQ2").End(xlDown)).Address & "," & Range("N3").Address(0, 0) & ")=1,1,"""")"         '(mark with 1)
        Range("ZI3").AutoFill Destination:=Range("ZI3:ZI" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZI3:ZI" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    Else
        Range("ZI3").Formula = "=if(countif(" & Range(Range("ZQ3"), Range("ZQ2").End(xlDown)).Address & "," & Range("N3").Address(0, 0) & ")<>1,1,"""")"         '(mark with 1)
        Range("ZI3").AutoFill Destination:=Range("ZI3:ZI" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZI3:ZI" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If

'    formnum = "A" & SherlockForm.CATNUMtxtform.Value
'    If SherlockForm.CatnumCheck.Value = False Then
'        desccnt = 0
'        Do
'            desccnt = desccnt + 1
'            findnmbr = "A" & Range("N2").Offset(desccnt, 0).Value
'            If findnmbr = formnum Then
'                Range("ZI2").Offset(desccnt, 0).Value = 1
'            End If
'        Loop Until IsEmpty(Range("A2").Offset(desccnt, 0))
'    Else
'        desccnt = 0
'        Do
'            desccnt = desccnt + 1
'            findnmbr = "A" & Range("N2").Offset(desccnt, 0).Value
'            If findnmbr <> formnum Then
'                Range("ZI2").Offset(desccnt, 0).Value = 1
'            End If
'        Loop Until IsEmpty(Range("A2").Offset(desccnt, 0))
'    End If
'
    '(Should it be included) Catnum and / or
    '-------------------------------
    If SherlockForm.CATNUMand = True Then
        Range("ZJ3").Formula = "=if(ZI3=1," & Chr(34) & "y" & Chr(34) & "," & Chr(34) & "x" & Chr(34) & ")"                                                             '(mark with 1)
        Range("ZJ3").AutoFill Destination:=Range("ZJ3:ZJ" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZJ3:ZJ" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    ElseIf SherlockForm.CATNUMdna <> True Then
        Range("ZJ3").Formula = "=if(ZI3=1," & Chr(34) & "y" & Chr(34) & ","""")"                                                            '(mark with 1)
        Range("ZJ3").AutoFill Destination:=Range("ZJ3:ZJ" & Cells(Rows.Count, "N").End(xlUp).Row)
        Range("ZJ3:ZJ" & Cells(Rows.Count, "N").End(xlUp).Row).Calculate
    End If
End If

'Which results will be included
'-------------------------------
If SherlockForm.IncludeAll.Value = True Then
    incallcnt = -1
    Do
    incallcnt = incallcnt + 1
    If (Range("ZB3").Offset(incallcnt, 0).Value <> "x" And Range("ZD3").Offset(incallcnt, 0).Value <> "x" And Range("ZF3").Offset(incallcnt, 0).Value <> "x" And Range("ZH3").Offset(incallcnt, 0).Value <> "x" And Range("ZJ3").Offset(incallcnt, 0).Value <> "x") And (Range("ZB3").Offset(incallcnt, 0).Value = "y" Or Range("ZD3").Offset(incallcnt, 0).Value = "y" Or Range("ZF3").Offset(incallcnt, 0).Value = "y" Or Range("ZH3").Offset(incallcnt, 0).Value = "y" Or Range("ZJ3").Offset(incallcnt, 0).Value = "y") Then
          Range("ZL3").Offset(incallcnt, 0).Value = 1
    End If
    Loop Until IsEmpty(Range("A3").Offset(incallcnt, 0))
Else
    incallcnt = -1
    Do
        incallcnt = incallcnt + 1
        If Range("ZK3").Offset(incallcnt, 0).Value <> 1 And (Range("ZB3").Offset(incallcnt, 0).Value <> "x" And Range("ZD3").Offset(incallcnt, 0).Value <> "x" And Range("ZF3").Offset(incallcnt, 0).Value <> "x" And Range("ZH3").Offset(incallcnt, 0).Value <> "x" And Range("ZJ3").Offset(incallcnt, 0).Value <> "x") And (Range("ZB3").Offset(incallcnt, 0).Value = "y" Or Range("ZD3").Offset(incallcnt, 0).Value = "y" Or Range("ZF3").Offset(incallcnt, 0).Value = "y" Or Range("ZH3").Offset(incallcnt, 0).Value = "y" Or Range("ZJ3").Offset(incallcnt, 0).Value = "y") Then
            Range("ZL3").Offset(incallcnt, 0).Value = 1
        End If
    Loop Until IsEmpty(Range("A3").Offset(incallcnt, 0))
End If


Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


End Sub
Sub PFfind()

Application.ScreenUpdating = False

If Not (SherlockForm.PFkeyword.Text = "" Or SherlockForm.PFkeyword.Text = " " Or Not SherlockForm.PFkeyword.Text <> "  ") Or Not (SherlockForm.PFcatnum.Value = "" Or SherlockForm.PFcatnum.Value = " " Or SherlockForm.PFcatnum.Value = "  ") Then

    'setup clinical tab if not already setup
    '-------------------------------------------
    Dim Sh As Worksheet, flg As Boolean
    For Each Sh In Worksheets
        If Sh.Name Like "ClinicalQC*" Then flg = True: Exit For
    Next
    
    If flg = True Then
        Sheets("Line Item Data").Range("ZA1").Value = 1
        Sheets("clinicalqc").Visible = True
        Sheets("ClinicalQC").Select
    Else
        Sheets("Line Item Data").Range("ZA1").Value = 1
        Sheets.Add After:=Sheets("Line Item Data")
        ActiveSheet.Name = "ClinicalQC"
        Range("A1").Value = "Original Sort"
        Range("B1").Value = "PSC"
        Range("C1").Value = "PIM key"
        Range("D1").Value = "Description"
        Range("E1").Value = "Manufacturer"
        Range("F1").Value = "Standard Manufacturer"
        Range("G1").Value = "Catalog Number"
        Range("H:H").ColumnWidth = 3
        Range("I1").Value = "Pricefile catalog number"
        Range("J1").Value = "Pricefile Description"
        Range("K1").Value = "Supplier Name"
        Rows("1:1").Interior.ColorIndex = 15
        Range("H:H").Interior.ColorIndex = 1
    End If
    
    'reset clinical values
    '-------------------------------------------
    Range("I2:K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear

    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets

    If InStr(sht.Name, "Pricing") > 0 And Not InStr(sht.Name, "Pricing") = 1 Then
        SuppNm = Left(sht.Name, Len(sht.Name) - 8) 'Sheets("Notes").Range("AA1").Offset(PFfindcnt, 0).Value
        sht.Visible = True
        sht.Select
        counter = counter + 1
        Application.StatusBar = "Supplier (" & counter & ")"

        Range("KD2:KD100000").ClearContents
        If sht.FilterMode = True Then
            Rows("1:1").AutoFilter
        End If
        
        'Catnum
        '-------------------------------
        If Not SherlockForm.PFcatnum.Value = "" And Not SherlockForm.PFcatnum.Value = " " And Not SherlockForm.PFcatnum.Value = "  " Then
            Range("KD2").Formula = "=if(countif('Line Item Data'!" & Range(Sheets("Line Item Data").Range("ZS3"), Sheets("Line Item Data").Range("ZS2").End(xlDown)).Address & "," & Range("A2").Address(0, 0) & ")=1,1,"""")"         '(mark with 1)
            Range("KD2").AutoFill Destination:=Range("KD2:KD" & Cells(Rows.Count, "C").End(xlUp).Row)
            Range("KD2:KD" & Cells(Rows.Count, "C").End(xlUp).Row).Calculate


'            formnum = "A" & SherlockForm.PFcatnum.Value
'            desccnt = 0
'            Do
'                desccnt = desccnt + 1
'                findnmbr = "A" & Range("A2").Offset(desccnt, 0).Value
'                If findnmbr = formnum Then
'                    Range("KD2").Offset(desccnt, 0).Value = 1
'                End If
'            Loop Until IsEmpty(Range("A2").Offset(desccnt, 0))
        End If

        'Keyword
        '-------------------------------
        If Not SherlockForm.PFkeyword.Text = "" And Not SherlockForm.PFkeyword.Text = " " And Not SherlockForm.PFkeyword.Text = "  " Then
            DescCnt = 0
            Do
                DescCnt = DescCnt + 1
                DoEvents
                Application.StatusBar = "Checking Pricefile keywords: row " & Range("A1").End(xlDown).Row - (Range("C1").Offset(DescCnt, 0).Row)
                innercnt = 0
                'mtchcnt = 0
                For Each c In Range(Sheets("Line Item Data").Range("ZR3"), Sheets("Line Item Data").Range("ZR2").End(xlDown))
                    innercnt = innercnt + 1
                    'Debug.Print LCase(Range("H2").Offset(desccnt, 0).Value)
                    'Debug.Print LCase(Range("ZO2").Offset(innercnt, 0).Value)
                    If InStr(LCase(Range("C1").Offset(DescCnt, 0).Value), LCase(Sheets("Line Item Data").Range("ZR2").Offset(innercnt, 0).Value)) = 0 Then
                        'mtchcnt = mtchcnt + 1
                        GoTo 5
                    End If
                Next
                'If mtchcnt = Application.CountA(Range(Range("ZO3"), Range("ZO2").End(xlDown))) Then
                    'Range("ZE2").Offset(desccnt, 0).Value = 1
                'End If
                Range("KD1").Offset(DescCnt, 0).Value = 1
5           Loop Until IsEmpty(Range("C1").Offset(DescCnt, 0))


'            desccnt = 0
'            'Debug.Print LCase(SherlockForm.PFkeyword.Text)
'            Do
'                desccnt = desccnt + 1
'                'Debug.Print LCase(Range("C3").Offset(desccnt, 0).Value)
'                'Debug.Print InStr(LCase(Range("C3").Offset(desccnt, 0).Value), LCase(SherlockForm.PFkeyword.Text))
'                If InStr(LCase(Range("C3").Offset(desccnt, 0).Value), LCase(SherlockForm.PFkeyword.Text)) > 0 Then
'                    Range("KD3").Offset(desccnt, 0).Value = 1
'                End If
'            Loop Until IsEmpty(Range("C3").Offset(desccnt, 0))
        End If
        

        'Count total rows
        '-------------------------------------------
        Range("C2").Select
        Range(Selection, Selection.End(xlDown)).Select
        ttlrows = Application.WorksheetFunction.CountA(Selection) - 1

        'Return results (filter)
        '-------------------------------------------
        If Not Application.WorksheetFunction.sum(Range("KD:KD")) = 0 Then
            On Error GoTo errhndlefilter
            ActiveSheet.Range(Range("A1"), Range("KK2").Offset(ttlrows, 0)).AutoFilter Field:=290, Criteria1:="<>"
            On Error GoTo 1
            Range(Range("A2"), Range("A2").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
            On Error GoTo 0
            Sheets("clinicalqc").Select
            Range("I1").Select
            If Not IsEmpty(ActiveCell.Offset(1, 0)) Then
                ActiveCell.End(xlDown).Offset(1, 0).Select
            Else
                ActiveCell.Offset(1, 0).Select
            End If
            ActiveSheet.Paste
            sht.Select
            Range(Range("C2"), Range("C2").Offset(ttlrows, 0)).SpecialCells(xlCellTypeVisible).Copy
            Sheets("clinicalqc").Select
            Range("J1").Select
            If Not IsEmpty(ActiveCell.Offset(1, 0)) Then
                ActiveCell.End(xlDown).Offset(1, 0).Select
            Else
                ActiveCell.Offset(1, 0).Select
            End If
            ActiveSheet.Paste
            Range("K1").Select
            If Not IsEmpty(ActiveCell.Offset(1, 0)) Then
                ActiveCell.End(xlDown).Offset(1, 0).Select
            Else
                ActiveCell.Offset(1, 0).Select
            End If
            Range(Selection, Range("J1").End(xlDown).Offset(0, 1)).Value = SuppNm
        End If
    End If
1 Next

SherlockForm.PFkeyword.Text = ""
SherlockForm.PFcatnum.Value = ""

'Debug.Print SherlockForm.PFkeyword.Text
'Debug.Print SherlockForm.PFcatnum.Value

Application.StatusBar = False

Exit Sub

errhndlefilter:
Rows("1:1").AutoFilter
On Error GoTo 0
Resume

End If

End Sub

Function ClinicalOOS() '(control As IRibbonControl)

Usr = FUN_findUsr

On Error GoTo 1
If Sheets("clinicalQC") Is ActiveSheet Then
    On Error GoTo 0
    Application.ScreenUpdating = False
    
    Set isect = Application.Intersect(ActiveCell, Range("G:G"))
    If isect Is Nothing Then
        MsgBox "Please select the catalog numbers you wish to assign."
    Else
        mtchitems = WorksheetFunction.CountA(Selection)
        Set mtchset = Selection.Offset(0, -6)
        Selection.Interior.Color = RGB(180, 0, 0)
        For Each c In mtchset
            If Not LCase(Usr) = "mlemay" Then
                Sheets("Line Item Data").Range("A:A").Find(what:=c.Value, lookat:=xlWhole).Offset(0, 13).Interior.Color = RGB(180, 0, 0)
            Else
                Sheets("Line Item Data").Range("A:A").Find(what:=c.Value, lookat:=xlWhole).Offset(0, 13).Interior.Color = RGB(254, 0, 0)
            End If
        Next
    End If

Else
1   If Not LCase(Usr) = "mlemay" Then
        Selection.Interior.Color = RGB(180, 0, 0)       '<<OUT OF SCOPE>>
    Else
        Selection.Interior.Color = RGB(254, 0, 0)
    End If
End If



End Function
Function ClinicalIS() '(control As IRibbonControl)

Usr = FUN_findUsr

On Error GoTo 1
If Sheets("clinicalQC") Is ActiveSheet Then
    On Error GoTo 0
    Application.ScreenUpdating = False
    
    Set isect = Application.Intersect(ActiveCell, Range("G:G"))
    If isect Is Nothing Then
        MsgBox "Please select the catalog numbers you wish to assign."
    Else
        mtchitems = WorksheetFunction.CountA(Selection)
        Set mtchset = Selection.Offset(0, -6)
        Selection.Interior.Color = 65280
        For Each c In mtchset
            If Not LCase(Usr) = "mlemay" Then
                Sheets("Line Item Data").Range("A:A").Find(what:=c.Value, lookat:=xlWhole).Offset(0, 7).Interior.Color = 65280
            Else
                Sheets("Line Item Data").Range("A:A").Find(what:=c.Value, lookat:=xlWhole).Offset(0, 13).Interior.Color = RGB(0, 127, 0)
            End If
        Next
    End If

Else
1   If Not LCase(Usr) = "mlemay" Then
        Selection.Interior.Color = 65280       '<<OUT OF SCOPE>>
    Else
        Selection.Interior.Color = RGB(0, 127, 0)
    End If
End If

End Function
Function ClinicalTBD() '(control As IRibbonControl)

Usr = FUN_findUsr

On Error GoTo 1
If Sheets("clinicalQC") Is ActiveSheet Then
    On Error GoTo 0
    Application.ScreenUpdating = False
    
    Set isect = Application.Intersect(ActiveCell, Range("G:G"))
    If isect Is Nothing Then
        MsgBox "Please select the catalog numbers you wish to assign."
    Else
        mtchitems = WorksheetFunction.CountA(Selection)
        Set mtchset = Selection.Offset(0, -6)
        Selection.Interior.Color = 652804
        For Each c In mtchset
            If Not LCase(Usr) = "mlemay" Then
                Sheets("Line Item Data").Range("A:A").Find(what:=c.Value, lookat:=xlWhole).Offset(0, 7).Interior.Color = 65280
            Else
                Sheets("Line Item Data").Range("A:A").Find(what:=c.Value, lookat:=xlWhole).Offset(0, 13).Interior.Color = RGB(254, 254, 0)
            End If
        Next
    End If

Else
1   If Not LCase(Usr) = "mlemay" Then
        Selection.Interior.Color = 652804       '<<OUT OF SCOPE>>
    Else
        Selection.Interior.Color = RGB(254, 254, 0)
    End If
End If



End Function
Sub SherlockRunKeys()
    SherlockPhraseDensity 1, "M"
    SherlockPhraseDensity 2, "O"
    SherlockPhraseDensity 3, "Q"
End Sub
Sub SherlockPhraseDensity(nWds As Long, col As Variant)
    Dim astr()      As String
    Dim i           As Long
    Dim j           As Long
    Dim cell        As Range
    Dim sPair       As String
    Dim rOut        As Range

    With CreateObject("Scripting.Dictionary")
        .CompareMode = vbTextCompare
        For Each cell In Range("D1", Cells(Rows.Count, "D").End(xlUp))
            astr = Split(Letters(cell.Value), " ")

            For i = 0 To UBound(astr) - nWds + 1
                sPair = vbNullString
                For j = i To i + nWds - 1
                    sPair = sPair & astr(j) & " "
                Next j
                sPair = Left(sPair, Len(sPair) - 1)

                If Not .exists(sPair) Then
                    .Add sPair, 1
                Else
                    .Item(sPair) = .Item(sPair) + 1
                End If
            Next i
        Next cell

        Set rOut = Columns(col).Resize(.Count, 2).Cells
        rOut.EntireColumn.ClearContents

        rOut.Columns(1).Value = Application.Transpose(.keys)
        rOut.Columns(2).Value = Application.Transpose(.items)

        rOut.Sort Key1:=rOut(1, 2), Order1:=xlDescending, _
                  Key2:=rOut(1, 1), Order1:=xlAscending, _
                  MatchCase:=False, Orientation:=xlTopToBottom, Header:=xlNo
        rOut.EntireColumn.AutoFit
    End With
End Sub
Function Letters(s As String) As String
    Dim i           As Long

    For i = 1 To Len(s)
        Select Case Mid(s, i, 1)
            Case "A" To "ÿ", "a" To "ÿ", "A" To "Z", "a" To "z", "0" To "9", "'"
                Letters = Letters & Mid(s, i, 1)
            Case Else
                Letters = Letters & " "
        End Select
    Next i
    Letters = WorksheetFunction.Trim(Letters)
End Function
Sub Clear()

    Columns("A:G").Select
    Range("G1").Activate
    Selection.ClearContents
    Range("A1").Select

End Sub
Sub METH_ClinicalClean(Optional LID As Boolean)


    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset

    If Not MainCall = 1 Then
        If Not FUN_Save = vbYes Then Exit Sub
        SetupSwitch = FUN_SetupSwitch
    End If
    
    If LID = True Then
        spendsheet = "Line Item Data"
        Mrkcol = Left(Sheets("line item data").Range("A4").End(xlToRight).Offset(0, 1).Address(0, 0), 2)
        FstRw = 5
    Else
        spendsheet = "Spend Search"
        Mrkcol = "AL"
        FstRw = 2
    End If
    Call FUN_TestForSheet("Items Removed")
    
    'find scopeguide
    '-------------------------------
    On Error GoTo errhndlNOSCOPE
    Sheets("Scopeguide").Visible = True
    On Error GoTo 0
    If Not Trim(Sheets("scopeguide").Range("A3").Value) = "" Then Set ISrng = Range(Sheets("scopeguide").Range("A3"), Sheets("scopeguide").Range("E2").End(xlDown).Offset(0, -4))
    If Not Trim(Sheets("scopeguide").Range("G3").Value) = "" Then Set osrng = Range(Sheets("scopeguide").Range("G3"), Sheets("scopeguide").Range("K2").End(xlDown).Offset(0, -4))
        
    'find active contracts
    '-------------------------------
    Call FUN_TestForSheet("xxCalculations")
    Cells.Clear
    Range("A1:A2").Value = "x"
    For i = 0 To ZeusForm.asscContracts.ListCount - 1
        Range("A1").End(xlDown).Offset(1, 0).Value = ZeusForm.asscContracts.List(i)
    Next
    
    'find expired contracts
    '-------------------------------
    'SQLstr = "SELECT CONTRACT_NUMBER FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con INNER JOIN OCSDW_GPO_CONTRACT gpo ON con.CONTRACT_ID = gpo.CONTRACT_ID AND gpo.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "' WHERE con.ATTRIBUTE_VALUE_NAME = '" & pscVar & "'"    'and con.STATUS_KEY = 'EXPIRED' AND NOT UPPER(con.CONTRACT_NAME) LIKE '%NOT BID%' AND con.CONTRACT_NUMBER LIKE '[a-z][a-z][0-9]%[0-9]' ORDER BY CONTRACT_NUMBER"
    strSELECT = "SELECT rtrim(con.CONTRACT_NUMBER)"
    strFROM = " FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con"
    strCMPY = " INNER JOIN OCSDW_GPO_CONTRACT gpo ON con.CONTRACT_ID = gpo.CONTRACT_ID AND gpo.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "'"
    strWHERE = " WHERE con.ATTRIBUTE_VALUE_NAME = '" & PSCVar & "' AND con.EXPORT_TYPE_KEY = 'M'"  'AND con.STATUS_KEY IN ('ACTIVE','SIGNED','PENDING','EXPIRED','TERMINATED')
    sqlstr = strSELECT & strFROM & strCMPY & strWHERE
    On Error Resume Next
    Set conn = Nothing
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

    Range("A1").End(xlDown).Offset(1, 0).CopyFromRecordset recset
        
    Set recset = Nothing
    Set conn = Nothing
    On Error GoTo 0
    
    Set conrng = Range(Range("A3"), Range("A1").End(xlDown))
    
    Sheets(spendsheet).Select
    lastrw = Range("A" & FstRw - 1).End(xlDown).Row
    Set PIMrng = Range("E" & FstRw & ":E" & lastrw)
    If LID = True Then
        PIMrng.ClearFormats
        PIMrng.Offset(0, -2).ClearFormats
    Else
        Cells.ClearFormats
    End If
    
    'On/Off contract
    '----------------------
    If Trim(Sheets("items removed").Range("A2").Value) = "" Then
        lstrmv = 1
    Else
        lstrmv = Sheets("items removed").Range("A1").End(xlDown).Row
    End If
    On Error GoTo errhndlNoMoreMtchs
    fndstrt = 1
    Do
        Application.StatusBar = "Removing off contract: " & fndstrt
        fndstrt = Range("F" & fndstrt & ":F" & lastrw).Find(what:="M", lookat:=xlWhole).Row
        
        If Application.CountIf(conrng, Trim(Range("H" & fndstrt).Value)) > 0 Then
            Range("X" & fndstrt).Interior.Color = RGB(0, 127, 0)    'Mark in (green)
        Else
            lstrmv = lstrmv + 1
            Range(Sheets("items removed").Range("A" & lstrmv), Sheets("items removed").Range("AK" & lstrmv)).Value = Range(Range("A" & fndstrt), Range("AK" & fndstrt)).Value
            Range(Mrkcol & fndstrt).Value = 1
        End If
    Loop Until fndstrt = Int(Mid(Application.StatusBar, InStr(Application.StatusBar, ": ") + 2, Len(Application.StatusBar)))
    
NoMoreMtchs:
    'PIM
    '------------------
    'If IsEmpty(isrng) And IsEmpty(osrng) Then
    Dim PIMArray(1 To 3) As Integer
    PIMArray(1) = 0
    PIMArray(2) = -2
    PIMArray(3) = -4
    
    For i = 1 To UBound(PIMArray)
        On Error GoTo errhndlNoMorePIM
        fndstrt = 1
        Do
            Application.StatusBar = "Finding unmatched PIMs: " & fndstrt
            fndstrt = Range("E" & fndstrt & ":E" & lastrw).Find(what:=PIMArray(i), lookat:=xlWhole).Row
            Range("E" & fndstrt).Interior.Color = 65535
        Loop Until fndstrt = Int(Mid(Application.StatusBar, InStr(Application.StatusBar, ": ") + 2, Len(Application.StatusBar)))

NoMorePIM:
    Next
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'    Else
        
'        Range(Mrkcol & fstrw).Formula = "=if('ScopeGuide'!" & ISrng.Address & ",E" & fstrw & ")"
'        Range(Mrkcol & fstrw).AutoFill Destination:=Range(Mrkcol & fstrw & ":" & Mrkcol & lastrw)
'        Range(Mrkcol & fstrw & ":" & Mrkcol & lastrw).Calculate
'        Range(Mrkcol & ":" & Mrkcol).SpecialCells(xlCellTypeConstants, 4).Offset(0, -354).Interior.Color = 65280
'
        
        On Error Resume Next
        For Each c In PIMrng
            DoEvents
            Application.StatusBar = "Finding in/out of scope: " & c.Row & " of " & lastrw
            'If c.Value = -2 Or c.Value = -4 Then
                'c.Interior.Color = 65535

            'Scopeguide
            '------------------
            If Application.CountIf(ISrng, c.Value) > 0 Then
                c.Interior.Color = 65280
            ElseIf Application.CountIf(osrng, c.Value) > 0 Then
                If Not IsEmpty(osrng) Then c.Interior.Color = RGB(180, 0, 0)
            End If
        Next
        On Error GoTo 0
    'End If
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'    On Error Resume Next
'    If Not IsEmpty(osrng) Then
'        TtlOS = osrng.Count
'        For Each c In osrng
'            Application.StatusBar = "Finding out of scope: " & c.Row & " of " & TtlOS
'            PIMrng.Replace what:=c.Value, Replacement:="x", lookat:=xlWhole
'            PIMrng.SpecialCells(xlCellTypeConstants, 2).Interior.Color = RGB(180, 0, 0)
'            PIMrng.SpecialCells(xlCellTypeConstants, 2).Value = c.Value
'        Next
'    End If
'    If Not IsEmpty(isrng) Then
'        TtlIS = isrng.Count
'        For Each c In isrng
'            Application.StatusBar = "Finding in scope: " & c.Row & " of " & TtlIS
'            PIMrng.Replace what:=c.Value, Replacement:="x", lookat:=xlWhole
'            PIMrng.SpecialCells(xlCellTypeConstants, 2).Interior.Color = 65280
'            PIMrng.SpecialCells(xlCellTypeConstants, 2).Value = c.Value
'        Next
'    End If
'    On Error GoTo 0


    'PSC
    '------------------
    On Error Resume Next
    'If Not Trim(Range("A3").Value) = "" Then    '<--check if only one item
        PIMrng.Offset(0, -2).Replace what:=PSCVar, replacement:="1", lookat:=xlWhole
        PIMrng.Offset(0, -2).SpecialCells(xlCellTypeConstants, 1).Interior.Color = 65280
        PIMrng.Offset(0, -2).SpecialCells(xlCellTypeConstants, 1).Value = PSCVar
    'ElseIf PIMrng.Offset(0, -2).Value = PSCVar Then
        'PIMrng.Offset(0, 2).Interior.Color = 65280
    'End If
    On Error GoTo 0
    
    'Remove and Clean
    '------------------
    On Error GoTo errhndlNOoos
    Application.StatusBar = "Removing out of scope...Please Wait"
    Range(Mrkcol & 1).ClearContents
    If Application.CountA(Range(Mrkcol & ":" & Mrkcol)) > 0 Then
        frstrw = Sheets("items removed").Range("A:A").Find(what:=Range("A" & Range(Mrkcol & 1).End(xlDown).Row).Value).Row
        Range(Mrkcol & ":" & Mrkcol).EntireColumn.SpecialCells(xlCellTypeConstants, 1).EntireRow.Delete  'Select
        'Sheets("items removed").Range("AA1").AutoFill Destination:=Sheets("items removed").Range("AA1:AA" & lstrmv)
        Sheets("items removed").Range("A2:A" & lstrmv).EntireRow.WrapText = False
        Sheets("items removed").Range("H" & frstrw & ":H" & lstrmv).Interior.ColorIndex = 3
    End If
    
NOoos:
    Sheets(spendsheet).Select
    Range("X2").Select
    Range("A:E").EntireColumn.Hidden = False
    Sheets("xxCalculations").Visible = False
    Application.StatusBar = False
    

Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNOSCOPE:
Call Import_Scopeguide
Resume Next

errhndlNOoos:
On Error GoTo 0
Resume NOoos

errhndlNoMoreMtchs:
On Error GoTo 0
Resume NoMoreMtchs

errhndlNoMorePIM:
On Error GoTo 0
Resume NoMorePIM



End Sub



