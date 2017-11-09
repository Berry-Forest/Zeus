Attribute VB_Name = "H__Members"
Sub RefreshMembers(ReportStatus As Integer)

    Dim DataRng As Range
    
    'Clear Member List
    '-------------------------
    For i = 0 To ZeusForm.asscMembers.ListCount - 1
        ZeusForm.asscMembers.RemoveItem (0)
    Next

    'Import Stdzn and add Mbrs
    '===================================================================================================================================================
    Call Import_Mbrs_And_Stdzn

    'Get Members
    '===================================================================================================================================================
    If ReportStatus > 1 Then
        If CreateReport = True Then
            Set DataRng = Sheets("Spend Search").Range("P2:P" & FUN_lastrow("P", "Spend Search"))
        Else
            Set DataRng = Sheets("Line item data").Range("P5:P" & FUN_lastrow("P", "Line Item Data"))
        End If
        If NetNm = "CAHN" Then Call Get_Members_From_Data(True, DataRng)               '<--[OWNERS]
'            'Call GetOwners(False, DataRng)
'        ElseIf NetNm = "MNS" Then
'            'Call GetOwners(True, DataRng)
'        End If
    End If
    
    If ReportStatus = 3 Then
        
        'Add Members
        '===================================================================================================================================================
        Call AddMbrsToReport
        
        'Get date ranges
        '===================================================================================================================================================
        Call Import_Dates  '>>>>>>>>>>
        Sheets("Index").Select
    
        'CleanUp
        '===================================================================================================================================================
        Call SetBKMRKs  '>>>>>>>>>>
        Range(MbrBkmrk.Offset(1, 0), MbrBkmrk.End(xlDown)).Select
        Selection.WrapText = False
        Selection.HorizontalAlignment = xlLeft
        MSGraphBKMRK.EntireColumn.WrapText = False
        NonConBKMRK.EntireColumn.WrapText = False
        ConvBKMRK.EntireColumn.WrapText = False
        
        Call label_owner_items
    End If
    
    
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::



End Sub
Sub Import_Mbrs_And_Stdzn()

Dim recset As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim TempArray() As String


'ZeusForm.asscStartDate.Text = ""
'ZeusForm.asscEndDate.Text = ""
ZeusForm.AnnualizedChk = False
Application.ScreenUpdating = False
For Each Wb In Workbooks
    If Wb.FullName = Stdzn_Index_PATH & "\" & NetNm & ".xlsx" Then wbopen = 1
Next

Application.DisplayAlerts = False
'connDB.Open "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & Stdzn_Index_PATH & "\" & NetNm & ".xlsx"
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Stdzn_Index_PATH & "\" & NetNm & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
            
sqlstr = "SELECT DISTINCT [Std_MID], [Name_Rolled_up_to_in_Report], [Owner] FROM [" & StdTbl & "$] WHERE NOT int([Std_MID]) = '' AND ([Current_Source] = 'RDM' OR [Current_Source] = '" & StdyTbl & "') ORDER BY [Name_Rolled_up_to_in_Report]"

On Error GoTo ERR_NOmbrREC
'On Error GoTo 0
recset.Open sqlstr, conn, adOpenStatic, adLockReadOnly
recset.MoveFirst

On Error GoTo ERR_mbrdone
OwnrNmbr = 0
ReDim MbrNames(1 To 1)
ReDim MbrMIDArray(1 To 1)
ReDim OwnrMbrArray(1 To 1)
Do
    
    'Mbrs
    '==================================================
'    If recset.Fields(1) = "MedStar Health" Then
'        Debug.Print
'    End If
    
    'Get Names
    '--------------
    prntcount = prntcount + 1
    ReDim Preserve MbrNames(1 To prntcount)
    ReDim Preserve MbrMIDArray(1 To prntcount)
    prevnm = recset.Fields(1)
    trimnm = Trim(FUN_convMbr(recset.Fields(1)))
    MbrNames(prntcount) = trimnm
    ZeusForm.asscMembers.AddItem trimnm
    
    'Get Owners
    '--------------
    If recset.Fields.Count > 2 Then
        If Not Trim(recset.Fields(2)) = "" Then
    
            'Get Names
            '--------------
            OwnrNmbr = OwnrNmbr + 1
            ReDim Preserve OwnrMbrArray(1 To OwnrNmbr)
            OwnrMbrArray(OwnrNmbr) = Trim(recset.Fields(2)) & "|" & trimnm
    
        End If
    End If
    
    'Get MIDs
    '--------------
    ReDim TempArray(1 To 1)
    MbrCount = 0
    Do
        MbrCount = MbrCount + 1
        ReDim Preserve TempArray(1 To MbrCount)
        TempArray(MbrCount) = Trim(recset.Fields(0))
        recset.MoveNext
    Loop Until Not recset.Fields(1) = prevnm
    MbrMIDArray(prntcount) = TempArray
    
Loop Until recset.EOF

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
''all system
''-----------------
'Sqlstr = "SELECT DISTINCT [SYSTEM_ID], [Name_Rolled_Up_To_In_Report] FROM [" & NtwkNmArray(NetPos) & "$] WHERE [RDM_MID] = "" AND NOT Current_Source = 'Not_Included' ORDER BY [Name_Rolled_Up_To_In_Report]"
'On Error GoTo errhndlNOallsysREC
'adoRecSet.Open Sqlstr, connDB, adOpenStatic, adLockReadOnly
'adoRecSet.MoveFirst
'
'ReDim SystemNames(1 To adoRecSet.RecordCount, 1 To 2)
'For i = 1 To adoRecSet.RecordCount
'    trimNm = Trim(FUN_convMbr(adoRecSet.Fields(0)))
'    SystemNames(i, 1) = trimNm
'    SystemNames(i, 2) = Trim(adoRecSet.Fields(1))
'    ZeusForm.asscSystems.AddItem trimNm
'    adoRecSet.MoveNext
'Next
'allsysdone:
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

CleanUp:
On Error Resume Next
MbrNMBR = UBound(MbrNames)
Set recset = Nothing
Set conn = Nothing
If Not wbopen = 1 Then Workbooks(NetNm & ".xlsx").Close (False)

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_NOmbrREC:
Resume CleanUp

ERR_mbrdone:
MbrMIDArray(prntcount) = TempArray
Resume CleanUp


End Sub
Sub Get_Members_From_Data(DataOnly As Boolean, DataRng As Range)



    If DataOnly = True Then
        MbrNMBR = 0
        For i = 0 To ZeusForm.asscMembers.ListCount - 1
            ZeusForm.asscMembers.RemoveItem (0)
        Next
    End If
    
    'Add Mbrs
    '-------------------------
    For Each c In DataRng
        MbrNm = Trim(c.Value)
        AddMbr = True
        For i = 0 To ZeusForm.asscMembers.ListCount - 1
            If ZeusForm.asscMembers.List(i) = MbrNm Then AddMbr = False
        Next
        If AddMbr = True Then ZeusForm.asscMembers.AddItem MbrNm
    Next
    
    MbrNMBR = ZeusForm.asscSystems.ListCount + ZeusForm.asscMembers.ListCount
    


End Sub
Sub GetOwners(StdOwner As Boolean, DataRng As Range)


    OwnrNmbr = 0
    
    If StdOwner = True Then
        
        'Get owners from standard list
        '-------------------------
        ReDim OwnerNames(1 To 1)
        OwnerNames(1) = "MedStar"
        OwnrNmbr = 1
        
    Else
        
        'Get owners from data
        '-------------------------
        ReDim OwnerNames(1 To 1)
        For Each c In DataRng
            MbrNm = Trim(c.Value)
            If InStr(MbrNm, "-") > 0 Then
                OwnrNm = Trim(Left(MbrNm, InStr(MbrNm, "-") - 1))
                AddOwnr = True
                For j = 1 To OwnrNmbr
                    If OwnerNames(j) = OwnrNm Then AddOwnr = False
                Next
                If AddOwnr = True Then
                    'If UBound(OwnerNames) = 0 Then ReDim OwnerNames(1 To 1)
                    OwnrNmbr = OwnrNmbr + 1
                    ReDim Preserve OwnerNames(1 To OwnrNmbr)
                    OwnerNames(OwnrNmbr) = OwnrNm
                End If
            End If
        Next
        
    End If
    
        
End Sub
Sub AddMbrsToReport()


    If OwnrNmbr > 0 Then
        
        'Get unique owners
        '----------------------
        For i = 1 To OwnrNmbr
            If Not InStr(ownr_str, ":" & Trim(Split(OwnrMbrArray(i), "|")(0)) & ":") > 0 Then
                ownr_str = ownr_str & ":" & Trim(Split(OwnrMbrArray(i), "|")(0)) & ":"
            End If
        Next
        ownr_str = Mid(ownr_str, 2, Len(ownr_str) - 2)
        'Unq_Ownr = UBound(Split(ownr_str, ":"))
    End If
    
    
    NonOwnerNmbr = ZeusForm.asscSystems.ListCount + ZeusForm.asscMembers.ListCount
    Unq_Ownr = UBound(Split(ownr_str, "::")) + 1
    MbrNMBR = NonOwnerNmbr + Unq_Ownr


'Delete Rows
'===================================================================================================================================================
    
    'Delete rows in mbr table
    '-------------------------
    For Each shp In Sheets("Index").Shapes
        If shp.Type = 8 And Val(Replace(shp.Name, "Check Box ", "")) > 1 Then shp.Delete
    Next
    LastMbrRow = Sheets("Index").Cells.Find(what:="Analysis Scope", lookat:=xlWhole).Row - 4
    Sheets("index").Rows("27:" & LastMbrRow).EntireRow.Delete
    Range(MbrBkmrk.Offset(1, 0), MbrBkmrk.Offset(17, 0)).Value = ""
    
    'Delete rows in summary tables
    '-------------------------
    CurrMbrNmbr = Range(MSGraphBKMRK.Offset(1, 0), MSGraphBKMRK.End(xlDown).Offset(-1, 0)).Count
    Range(prsBKMRK.Offset(2, 0), prsBKMRK.Offset(CurrMbrNmbr, 0)).EntireRow.Delete
    Range(BenchBKMRK.Offset(2, 0), BenchBKMRK.Offset(CurrMbrNmbr, 0)).EntireRow.Delete
    Range(MSGraphBKMRK.Offset(2, 0), MSGraphBKMRK.Offset(CurrMbrNmbr, 0)).EntireRow.Delete

    For i = 1 To suppNMBR
        RowOffset = (CurrMbrNmbr + 8) * (suppNMBR - i)
        Range(NonConBKMRK.Offset(RowOffset + 2, 0), NonConBKMRK.Offset(RowOffset + CurrMbrNmbr, 0)).EntireRow.Delete
        Range(ConvBKMRK.Offset(RowOffset + 2, 0), ConvBKMRK.Offset(RowOffset + CurrMbrNmbr, 0)).EntireRow.Delete
    Next

'Add Rows
'===================================================================================================================================================
    
    'Add Rows to member table if needed
    '-------------------------
    If MbrNMBR > 17 Then
        Sheets("Index").Rows("26").EntireRow.Copy
        Sheets("Index").Rows("27:" & 27 + MbrNMBR - 18).Insert
    End If
    
    'add check boxes
    '-------------------------
    For i = 1 To MbrNMBR - 1
        Set ChkStart = MbrBkmrk.Offset(1 + i, -1)
        Set NewChkBx = Sheets("index").CheckBoxes.Add(ChkStart.Left + ChkStart.Width / 2 - 5, ChkStart.Top - ChkStart.Height / 32, 10, 10)
        NewChkBx.Caption = ""
        NewChkBx.LinkedCell = MbrBkmrk.Offset(1 + i, -7).Address
        NewChkBx.Value = True
    Next

    'Add rows in summary tables
    '-------------------------
    If MbrNMBR > 1 Then
    
        Set MSGraph = Sheets("initiative spend overview").Shapes("MS Graph")
        Set BenchGraph = Sheets("initiative spend overview").Shapes("Benchmark Graph")
        
        prsBKMRK.Offset(1, 0).EntireRow.Copy
        Range(prsBKMRK.Offset(2, 0), prsBKMRK.Offset(MbrNMBR, 0)).EntireRow.Insert
        BenchBKMRK.Offset(1, 0).EntireRow.Copy
        Range(BenchBKMRK.Offset(2, 0), BenchBKMRK.Offset(MbrNMBR, 0)).EntireRow.Insert
        MSGraphBKMRK.Offset(1, 0).EntireRow.Copy
        Range(MSGraphBKMRK.Offset(2, 0), MSGraphBKMRK.Offset(MbrNMBR, 0)).EntireRow.Insert
        
        For Each c In Sheets("initiative spend overview").Shapes
            If c.Name = BenchGraph.Name And Not c.ID = BenchGraph.ID Then
                c.Delete
            ElseIf c.Name = MSGraph.Name And Not c.ID = MSGraph.ID Then
                c.Delete
            End If
        Next
        
        For i = 1 To suppNMBR
            RowOffset = (MbrNMBR + 8) * (i - 1)
            NonConBKMRK.Offset(RowOffset + 1, 0).EntireRow.Copy
            Range(NonConBKMRK.Offset(RowOffset + 2, 0), NonConBKMRK.Offset(RowOffset + MbrNMBR, 0)).EntireRow.Insert
        Next
        
        For i = 1 To suppNMBR
            RowOffset = (MbrNMBR + 8) * (i - 1)
            ConvBKMRK.Offset(RowOffset + 1, 0).EntireRow.Copy
            Range(ConvBKMRK.Offset(RowOffset + 2, 0), ConvBKMRK.Offset(RowOffset + MbrNMBR, 0)).EntireRow.Insert
        Next
    
    End If
    
'Adjust Borders
'===================================================================================================================================================
    Range(MSGraphBKMRK.Offset(1, 0), MSGraphBKMRK.Offset(MbrNMBR, 0)).Borders(xlInsideHorizontal).LineStyle = xlNone
    Range(BenchBKMRK.Offset(1, 0), BenchBKMRK.Offset(MbrNMBR, 0)).Borders(xlInsideHorizontal).LineStyle = xlNone
    Range(prsBKMRK.Offset(1, 0), prsBKMRK.Offset(MbrNMBR, 0)).Borders(xlInsideHorizontal).LineStyle = xlNone
    
    For i = 1 To suppNMBR
        RowOffset = (MbrNMBR + 8) * (i - 1)
        Range(NonConBKMRK.Offset(RowOffset + 1, 0), NonConBKMRK.Offset(RowOffset + MbrNMBR, 0)).Borders(xlInsideHorizontal).LineStyle = xlNone
        Range(NonConBKMRK.Offset(RowOffset + 1, 0), NonConBKMRK.Offset(RowOffset + MbrNMBR, 0)).BorderAround Color:=12566463
        Range(ConvBKMRK.Offset(RowOffset + 1, 0), ConvBKMRK.Offset(RowOffset + MbrNMBR, 0)).Borders(xlInsideHorizontal).LineStyle = xlNone
        Range(ConvBKMRK.Offset(RowOffset + 1, 0), ConvBKMRK.Offset(RowOffset + MbrNMBR, 0)).BorderAround Color:=12566463
    Next
    
'Adjust formulas
'===================================================================================================================================================
        
    'Adjust owner formulas
    '-------------------------
    'If NetNm = "CAHN" Or NetNm = "MNS" Then
        If OwnrNmbr > 0 Then
            
            If Application.CountIf(Sheets("line item data").Rows("4:4"), "Owners") > 0 Then
                OwnrCol = Sheets("line item data").Rows("4:4").Find(what:="Owners", lookat:=xlWhole).EntireColumn.Address
                OwnrCol = Left(OwnrCol, InStr(OwnrCol, ":") - 1)
            Else
                Sheets("line item data").Columns.Hidden = False
                OwnrCol = Sheets("line item data").Range("A4").End(xlToRight).Offset(0, 1).EntireColumn.Address
                OwnrCol = Left(OwnrCol, InStr(OwnrCol, ":") - 1)
            End If
            
            Range(MSGraphBKMRK.Offset(NonOwnerNmbr + 1, 0), MSGraphBKMRK.Offset(MbrNMBR, 0)).EntireRow.Replace what:="$P", replacement:=OwnrCol, lookat:=xlPart
            Range(BenchBKMRK.Offset(NonOwnerNmbr + 1, 0), BenchBKMRK.Offset(MbrNMBR, 0)).EntireRow.Replace what:="$P", replacement:=OwnrCol, lookat:=xlPart
            Range(prsBKMRK.Offset(NonOwnerNmbr + 1, 0), prsBKMRK.Offset(MbrNMBR, 0)).EntireRow.Replace what:="$P", replacement:=OwnrCol, lookat:=xlPart
            
            For i = 1 To suppNMBR
                RowOffset = (MbrNMBR + 8) * (i - 1)
                Range(NonConBKMRK.Offset(RowOffset + NonOwnerNmbr + 1, 0), NonConBKMRK.Offset(RowOffset + MbrNMBR, 0)).EntireRow.Replace what:="$P", replacement:=OwnrCol, lookat:=xlPart
                Range(ConvBKMRK.Offset(RowOffset + NonOwnerNmbr + 1, 0), ConvBKMRK.Offset(RowOffset + MbrNMBR, 0)).EntireRow.Replace what:="$P", replacement:=OwnrCol, lookat:=xlPart
                Range(ConvBKMRK.Offset(RowOffset + NonOwnerNmbr + 1, 0), ConvBKMRK.Offset(RowOffset + MbrNMBR, 0)).EntireRow.Replace what:="$Z", replacement:=Left(Range(OwnrCol & 1).Offset(0, 1).EntireColumn.Address, InStr(Range(OwnrCol & 1).Offset(0, 1).EntireColumn.Address, ":") - 1), lookat:=xlPart
            Next
                            
        End If
    'End If
    
    'Adjust totals formulas
    '-------------------------
    MSGraphBKMRK.End(xlDown).EntireRow.Replace what:=MSGraphBKMRK.Offset(1, 0).Row & ")", replacement:=MSGraphBKMRK.Offset(NonOwnerNmbr, 0).Row & ")", lookat:=xlPart
    BenchBKMRK.End(xlDown).EntireRow.Replace what:=BenchBKMRK.Offset(1, 0).Row & ")", replacement:=BenchBKMRK.Offset(NonOwnerNmbr, 0).Row & ")", lookat:=xlPart
    BenchBKMRK.Offset(-2, 1).Replace what:=":B" & BenchBKMRK.Offset(1, 0).Row, replacement:=":B" & BenchBKMRK.Offset(NonOwnerNmbr, 0).Row, lookat:=xlPart
    prsBKMRK.End(xlDown).EntireRow.Replace what:=prsBKMRK.Offset(1, 0).Row & ")", replacement:=prsBKMRK.Offset(NonOwnerNmbr, 0).Row & ")", lookat:=xlPart
    
    For i = 1 To suppNMBR
        RowOffset = (MbrNMBR + 8) * (i - 1)
        NonConBKMRK.Offset(RowOffset + MbrNMBR + 1).EntireRow.Replace what:=NonConBKMRK.Offset(RowOffset + 1, 0).Row & ")", replacement:=NonConBKMRK.Offset(RowOffset + NonOwnerNmbr, 0).Row & ")", lookat:=xlPart
        ConvBKMRK.Offset(RowOffset + MbrNMBR + 1).EntireRow.Replace what:=ConvBKMRK.Offset(RowOffset + 1, 0).Row & ")", replacement:=ConvBKMRK.Offset(RowOffset + NonOwnerNmbr, 0).Row & ")", lookat:=xlPart
        ConvBKMRK.Offset(RowOffset + MbrNMBR + 1).EntireRow.Replace what:="$" & ConvBKMRK.Offset(RowOffset + 1, 0).Row, replacement:="$" & ConvBKMRK.Offset(RowOffset + NonOwnerNmbr, 0).Row, lookat:=xlPart
    Next
    
    'Adjust member name formulas
    '-------------------------
    frstNm = NonConBKMRK.Row - 1
    For i = 1 To suppNMBR
        RowOffset = (MbrNMBR + 8) * (i - 1)
        For j = 2 To MbrNMBR
            NonConBKMRK.Offset(RowOffset + j, 0).Replace what:="$" & frstNm + 2, replacement:="$" & frstNm + j + 1, lookat:=xlPart
            NonConBKMRK.Offset(RowOffset + j, 0).Replace what:="$" & frstNm, replacement:="$" & frstNm + j - 1, lookat:=xlPart
            NonConBKMRK.Offset(RowOffset + j, -1).Replace what:="$" & frstNm, replacement:="$" & frstNm + j - 1, lookat:=xlPart
            ConvBKMRK.Offset(RowOffset + j, 0).Replace what:="$" & frstNm + 2, replacement:="$" & frstNm + j + 1, lookat:=xlPart
            ConvBKMRK.Offset(RowOffset + j, 0).Replace what:="$" & frstNm, replacement:="$" & frstNm + j - 1, lookat:=xlPart
            ConvBKMRK.Offset(RowOffset + j, -1).Replace what:="$" & frstNm, replacement:="$" & frstNm + j - 1, lookat:=xlPart
        Next
    Next
    
'Populate Mbrs
'===================================================================================================================================================
    
    'populate Mbrs from userform
    '-------------------------
    Sheets("Index").Select
    For i = 0 To ZeusForm.asscSystems.ListCount - 1
        MbrBkmrk.Offset(i + 1, 0).Value = Trim(ZeusForm.asscSystems.List(i))
    Next
    For i = 0 To ZeusForm.asscMembers.ListCount - 1
        MbrBkmrk.Offset(ZeusForm.asscSystems.ListCount + i + 1, 0).Value = Trim(ZeusForm.asscMembers.List(i))
    Next
    Call FUN_Sort("Index", Range(MbrBkmrk.Offset(1, 0), MbrBkmrk.Offset(NonOwnerNmbr, 0)), MbrBkmrk.Offset(1, 0), 1)
    
    'Populate Owners from Array
    '-------------------------
    If Unq_Ownr > 0 Then
        For i = 0 To Unq_Ownr - 1
            MbrBkmrk.End(xlDown).Offset(1, 0).Value = Split(ownr_str, "::")(i)
        Next
        Call FUN_Sort("Index", Range(MbrBkmrk.Offset(NonOwnerNmbr + 1, 0), MbrBkmrk.Offset(MbrNMBR, 0)), MbrBkmrk.Offset(NonOwnerNmbr + 1, 0), 1)
    End If
    
        
End Sub
Sub adjbxs()

        For Each c In ActiveSheet.Shapes
            If c.Type = 8 Then c.Left = Range("G9").Offset(i, 0).Left + Range("G9").Offset(i, 0).Width / 2 - 5
        Next

End Sub
'Sub AdjustTotals()
'
'
'    'MSgraph
'    '----------------------
'    For Each c In Range(MSGraphBKMRK.End(xlDown).Offset(0, 1), MSGraphBKMRK.End(xlDown).End(xlToRight)).Replace
'        c.Formula = Replace(c.Formula, ":" & MSGraphBKMRK.Offset(1, 0).Row, ":" & MSGraphBKMRK.Offset(nonownernmbr, 0).Row)
'        If OwnrNmbr > 0 Then c.Formula = Replace(c.Formula, c.Offset(-1, 0).Address(0, 0), c.Offset(-OwnrNmbr - 1, 0).Address(0, 0))
'    Next
'
'    'Benchmark
'    '----------------------
'    For Each c In Range(BenchBKMRK.End(xlDown).Offset(0, 1), BenchBKMRK.End(xlDown).Offset(0, 8))
'        c.Formula = Replace(c.Formula, c.Offset(-1, 0).Address(0, 0), c.Offset(-OwnrNmbr - 1, 0).Address(0, 0))
'    Next
'
'    'PRS
'    '----------------------
'    For Each c In Range(prsBKMRK.End(xlDown).Offset(0, 1), prsBKMRK.End(xlDown).End(xlToRight))
'        c.Formula = Replace(c.Formula, c.Offset(-1, 0).Address(0, 0), c.Offset(-OwnrNmbr - 1, 0).Address(0, 0))
'    Next
'
'    'Conversion/Non-Con
'    '----------------------
'    For supp = 1 To 10
'        For Each c In Range(NonConBKMRK.Offset((MbrNMBR + 1) + (MbrNMBR + 8) * (supp - 1), 1), NonConBKMRK.Offset((MbrNMBR + 1) + (MbrNMBR + 8) * (supp - 1), 0).End(xlToRight))
'            c.Formula = Replace(c.Formula, c.Offset(-1, 0).Address(0, 0), c.Offset(-OwnrNmbr - 1, 0).Address(0, 0))
'        Next
'        For Each c In Range(ConvBKMRK.Offset((MbrNMBR + 1) + (MbrNMBR + 8) * (supp - 1), 1), ConvBKMRK.Offset((MbrNMBR + 1) + (MbrNMBR + 8) * (supp - 1), 0).End(xlToRight))
'            c.Formula = Replace(c.Formula, c.Offset(-1, 0).Address(0, 0), c.Offset(-OwnrNmbr - 1, 0).Address(0, 0))
'        Next
'
'        'Fix unique product count formulas
'        '----------------------
'        For Each c In Range(ConvBKMRK.Offset((MbrNMBR - OwnrNmbr) + (MbrNMBR + 8) * (supp - 1), 8), ConvBKMRK.Offset(MbrNMBR + (MbrNMBR + 8) * (supp - 1), 9))
'            c.Formula = Replace(c.Formula, "$Z", "$MV")
'            c.Formula = Replace(c.Formula, "$P", "$MU")
'        Next
'    Next
'
'
'End Sub
Function CAHN_Function(StrtCell, OffsetStr)


Do
    OffRw = Left(OffsetStr, InStr(OffsetStr, ";") - 1)
    CAHN_Function = CAHN_Function & Range(StrtCell).Offset(OffRw, 0).Address(0, 0) & "+"
    OffsetStr = Mid(OffsetStr, InStr(OffsetStr, ";") + 1, Len(OffsetStr))
Loop Until InStr(OffsetStr, ";") = 0
CAHN_Function = "=" & Left(CAHN_Function, Len(CAHN_Function) - 1)


End Function
Sub label_owner_items()


        'Create Owners Column
        '-----------------------
        If Application.CountIf(Sheets("line item data").Rows("4:4"), "Owners") > 0 Then
            Set OwnrCol = Sheets("line item data").Rows("4:4").Find(what:="Owners", lookat:=xlWhole)
            Range(OwnrCol.Offset(1, 0), OwnrCol.Offset(ItmNmbr, 1)).ClearContents
        Else
            Sheets("line item data").Columns.Hidden = False
            Set OwnrCol = Sheets("line item data").Range("A4").End(xlToRight).Offset(0, 1)
            OwnrCol.Value = "Owners"
            OwnrCol.Font.Bold = True
            OwnrCol.Interior.Color = 65535
            OwnrCol.Offset(0, 1).Value = "Owner Unique Count"
            OwnrCol.Offset(0, 1).Font.Bold = True
            OwnrCol.Offset(0, 1).Interior.Color = 65535
        End If
        
        On Error GoTo ERR_NoOwner
        For i = 1 To ItmNmbr
            CurrRw = 4 + i
            MbrNm = Sheets("line item data").Range("P" & CurrRw).Value
            For j = 1 To OwnrNmbr
                If MbrNm = Trim(Split(OwnrMbrArray(j), "|")(1)) Then
                    OwnrCol.Offset(CurrRw - 4, 0).Value = Trim(Split(OwnrMbrArray(j), "|")(0))
                    OwnrCol.Offset(CurrRw - 4, 1).Formula = "=IF(SUMPRODUCT(($X$5:$X" & CurrRw & "=$X" & CurrRw & ")*($MU$5:$MU" & CurrRw & "=$MU" & CurrRw & "))>1,0,1)"
                    Exit For
                End If
            Next
            'OwnrCol.Offset(i, 0).Value = Trim(Left(Range("P" & CurrRw).Value, InStr(Range("P" & CurrRw).Value, "-") - 1))
NxtOwner:
        Next
        
        
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::
ERR_NoOwner:
OwnrCol.Offset(i, 0).Value = Range("P4").Offset(i, 0).Value
Resume NxtOwner
            
            
End Sub
