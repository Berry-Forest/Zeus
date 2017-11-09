Attribute VB_Name = "D__QC"
Public ManualCheck As Boolean
Public NoteUnhide As Boolean
Public NoteText As String
Public QCMain As Boolean
Sub QC_Main()

QCMain = True

For i = 1 To 49
    Call QCChecks(i)
Next

QCMain = False


End Sub
Sub QCChecks(CurrCheck)


    NoteUnhide = False
    Set StsBox = QCform.Controls("StatusBox" & CurrCheck)
    
    If CurrCheck = 1 Then
        CheckStatus = DATname
    ElseIf CurrCheck = 2 Then
        CheckStatus = VarRngs
    ElseIf CurrCheck = 3 Then
        CheckStatus = SpendTies_MS
    ElseIf CurrCheck = 4 Then
        CheckStatus = SpendTies_Bench
    ElseIf CurrCheck = 5 Then
        CheckStatus = SpendTies_PRS
    ElseIf CurrCheck = 6 Then
        CheckStatus = SpendTies_NC
    ElseIf CurrCheck = 7 Then
        CheckStatus = SpendTies_Conv
    ElseIf CurrCheck = 8 Then
        CheckStatus = SpendTies_LineItem
    ElseIf CurrCheck = 9 Then
        CheckStatus = ReportDate
    ElseIf CurrCheck = 10 Then
        CheckStatus = MbrDates
    ElseIf CurrCheck = 11 Then
        CheckStatus = MbrsCorrect
    ElseIf CurrCheck = 12 Then
        CheckStatus = UNSPSC
    ElseIf CurrCheck = 13 Then
        CheckStatus = TierInfo
    ElseIf CurrCheck = 14 Then
        CheckStatus = IOCcnt
    ElseIf CurrCheck = 15 Then
        CheckStatus = SpendReasonable
    ElseIf CurrCheck = 16 Then
        CheckStatus = PRS
    ElseIf CurrCheck = 17 Then
        CheckStatus = AotrsLessThan
    ElseIf CurrCheck = 18 Then
        CheckStatus = MSgraphData
    ElseIf CurrCheck = 19 Then
        CheckStatus = MSgraphReadable
    ElseIf CurrCheck = 20 Then
        CheckStatus = BenchGraphData
    ElseIf CurrCheck = 21 Then
        CheckStatus = BenchGraphReadable
    ElseIf CurrCheck = 22 Then
        CheckStatus = PropTotalsSum
    ElseIf CurrCheck = 23 Then
        CheckStatus = ConMftrMtchMPP
    ElseIf CurrCheck = 24 Then
        CheckStatus = NonConMftrNmsStd
    ElseIf CurrCheck = 25 Then
        CheckStatus = OneMftrPerCatnum
    ElseIf CurrCheck = 26 Then
        If StsBox.BackColor = &HFFFF& And Not QCMain = True Then
            CheckStatus = BlankMftrs(True)
        Else
            CheckStatus = BlankMftrs(False)
        End If
    ElseIf CurrCheck = 27 Then
        CheckStatus = CatnumsStdzd
    ElseIf CurrCheck = 28 Then
        CheckStatus = BlankCatnums
    ElseIf CurrCheck = 29 Then
        CheckStatus = PlvlVars
    ElseIf CurrCheck = 30 Then
        CheckStatus = MbrBenchVars
    ElseIf CurrCheck = 31 Then
        CheckStatus = SuppVars
    ElseIf CurrCheck = 32 Then
        CheckStatus = SuppBenchVars
    ElseIf CurrCheck = 33 Then
        CheckStatus = PkgMsmtch
    ElseIf CurrCheck = 34 Then
        If StsBox.BackColor = &HFFFF& And Not QCMain = True Then
            CheckStatus = OOSRmvd(True)
        Else
            CheckStatus = OOSRmvd(False)
        End If
    ElseIf CurrCheck = 35 Then
        CheckStatus = Novaplus
    ElseIf CurrCheck = 36 Then
        CheckStatus = DupCat
    ElseIf CurrCheck = 37 Then
        CheckStatus = BestPriceFrmla
    ElseIf CurrCheck = 38 Then
        CheckStatus = ZeroItmsRmvd
    ElseIf CurrCheck = 39 Then
        CheckStatus = UnqualTrsRmvd
    ElseIf CurrCheck = 40 Then
        CheckStatus = CurrPricingUsed
    ElseIf CurrCheck = 41 Then
        CheckStatus = XrefSorted
    ElseIf CurrCheck = 42 Then
        CheckStatus = XrefCleansed
    ElseIf CurrCheck = 43 Then
        CheckStatus = AdminFees
    ElseIf CurrCheck = 44 Then
        CheckStatus = BMPSorted
    ElseIf CurrCheck = 45 Then
        CheckStatus = FullyCalc
    ElseIf CurrCheck = 46 Then
        CheckStatus = FontsFormatted
    ElseIf CurrCheck = 47 Then
        CheckStatus = OrgnlsPosted
    ElseIf CurrCheck = 48 Then
        CheckStatus = DataIntact
    ElseIf CurrCheck = 49 Then
        CheckStatus = NoHardcoded
    End If

    'Update Status Box
    '--------------------
    If CheckStatus = 1 Then
        StsBox.BackColor = &HFF00&
    ElseIf CheckStatus = 2 Then
        StsBox.BackColor = &HFF&
    ElseIf CheckStatus = 3 Then
        StsBox.BackColor = &HFFFF&
    End If
    
    'hide/unhide note
    '--------------------
    If NoteUnhide = True Then
        QCform.Controls("Note" & CurrCheck).Visible = True
        If ManualCheck = True Then
            QCform.Controls("Note" & CurrCheck).Caption = "M " & NoteText
            QCform.Controls("Note" & CurrCheck).Width = 9
            QCform.Controls("Note" & CurrCheck).Left = StsBox.Left - 10
        Else
            QCform.Controls("Note" & CurrCheck).Caption = "! " & NoteText
            QCform.Controls("Note" & CurrCheck).Width = 4.5
            QCform.Controls("Note" & CurrCheck).Left = StsBox.Left - 4.5
        End If
        NoteUnhide = False
        ManualCheck = False
    ElseIf NoteUnhide = False Then
        QCform.Controls("Note" & CurrCheck).Visible = False
    End If


End Sub
Function DATname() As Integer

On Error GoTo ERR_noName

'<<Navigation>>
'------------
Call FUN_TestForSheet("Notes")
Sheets("Notes").Rows("1:1").Find(what:="DAT:").Offset(0, 1).Select

If Not Trim(ActiveCell.Value) = "" Then
    DATname = 1
Else
    DATname = 2
End If


Exit Function
':::::::::::::::
ERR_noName:
DATname = 2
Exit Function


End Function
Function VarRngs() As Integer

On Error GoTo ERR_noVarRng

'<<Navigation>>
'------------
Call FUN_TestForSheet("Notes")
Rows("2:2").Find(what:="PriceLeveling Variance Range:").Select

'Check
'------------
If Not Trim(ActiveCell.Offset(0, 1).Value) = "" And Not Trim(ActiveCell.Offset(1, 1).Value) = "" And Not Trim(ActiveCell.Offset(2, 1).Value) = "" Then
    VarRngs = 1
Else
    VarRngs = 2
End If


Exit Function
':::::::::::::::
ERR_noVarRng:
VarRngs = 2
Exit Function




End Function
Function SpendTies_MS() As Integer

On Error Resume Next

'<<Navigation>>
'--------------
Call FUN_TestForSheet("initiative Spend overview")
MSGraphBKMRK.Select

If Check_ttls("MSgraph") = True Then
    SpendTies_MS = 1
Else
    SpendTies_MS = 2
End If


End Function
Function SpendTies_Bench() As Integer

On Error Resume Next

'<<Navigation>>
'--------------
Call FUN_TestForSheet("initiative Spend overview")
BenchBKMRK.Select

If Check_ttls("Bench") = True Then
    SpendTies_Bench = 1
Else
    SpendTies_Bench = 2
End If

End Function
Function SpendTies_PRS() As Integer

On Error Resume Next

'<<Navigation>>
'--------------
Call FUN_TestForSheet("initiative Spend overview")
prsBKMRK.Select

If Check_ttls("PRS") = True Then
    SpendTies_PRS = 1
Else
    SpendTies_PRS = 2
End If


End Function
Function SpendTies_NC() As Integer

On Error Resume Next

'<<Navigation>>
'--------------
Call FUN_TestForSheet("Vizient Contracts - NC")
NonConBKMRK.Select

If Check_ttls("NC") = True Then
    SpendTies_NC = 1
Else
    SpendTies_NC = 2
End If

End Function
Function SpendTies_Conv() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Vizient Contracts - Conv")
ConvBKMRK.Select

'<<Check>>
'=========================================
If Check_ttls("Conv") = True Then
    SpendTies_Conv = 1
Else
    SpendTies_Conv = 2
End If

End Function
Function SpendTies_LineItem() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line Item data")
Range("AJ3").Select

'<<Check>>
'=========================================
If Check_ttls("LID") = True Then
    SpendTies_LineItem = 1
Else
    SpendTies_LineItem = 2
End If


End Function
Function ReportDate() As Integer


'<<Navigation>>
'=========================================
Call FUN_TestForSheet("index")
Range("C8").Select

'<<Check>>
'=========================================
If Not Range("C8").Value = "Date" Then
    ReportDate = 1
Else
    ReportDate = 2
End If


End Function
Function MbrDates() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("index")
Range("C:E").Find(what:="Primary").Offset(1, 0).Select

'<<Check>>
'=========================================
If Not ActiveCell.Value = "" Then
    MbrDates = 1
Else
    MbrDates = 2
End If


End Function
Function MbrsCorrect() As Integer


'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Notes")
Range("AH1").Select

'<<Check>>
'=========================================
Range("AH1").Value = "Standardized Names For Analysis"
unstdFLG = 0
If NetNm = "CAHN" Then
    Range("AH2").Value = "No standardized names to be included for CAHN.  Only members with spend to be included per PAT."
    Range("AH2").Interior.ColorIndex = 3
    MbrsCorrect = 1
Else
    If Trim(Range("AH2")) = "" Then

        'import systems from form
        '------------------------------
        For i = 0 To ZeusForm.asscSystems.ListCount - 1
            Range("AH2").Offset(i, 0).Value = ZeusForm.asscSystems.List(i)
        Next
        For i = 0 To ZeusForm.asscMembers.ListCount - 1
            Range("AH2").Offset(ZeusForm.asscSystems.ListCount + i, 0).Value = ZeusForm.asscMembers.List(i)
        Next
        
    End If
    
    Set ReportMbrs = Range(MbrBkmrk.Offset(1, 0), MbrBkmrk.Offset(MbrNMBR, 0))
    Set IDXmbrs = Range(Range("AH2"), Range("AH1").End(xlDown))
    
    For Each c In IDXmbrs
        If Not WorksheetFunction.CountIf(ReportMbrs, c.Value) > 0 Then
            c.Interior.Color = 16711935
            MbrsCorrect = 2
        ElseIf c.Interior.Color = 16711935 Then
            c.Interior.ColorIndex = 0
        End If
    Next
    If Not MbrsCorrect = 2 Then MbrsCorrect = 1

End If



End Function
Function UNSPSC() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("index")
Range("H:J").Find(what:="Product Coverage").Select

'<<Check>>
'=========================================
If Not ActiveCell.Offset(1, 0) = "" Then
    UNSPSC = 1
Else
    UNSPSC = 3
End If


End Function
Function TierInfo() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("index")
ConTblBKMRK.Select

'<<Check>>
'=========================================
QCFlg = True
QCChkFlg = True
Call Import_TierInfo
QCFlg = False

If QCChkFlg = True Then
    TierInfo = 1
Else
    TierInfo = 2
    NoteUnhide = True
    NoteText = "Tier info does not match MPP."
End If


End Function
Function IOCcnt() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("index")
ConTblBKMRK.Offset(7, 0).Select

'<<Check>>
'=========================================
IOCcnt = 1
For i = 1 To suppNMBR
    If Not Application.CountA(Sheets(FUN_SuppName(i) & " Pricing").Range("A:A")) - 1 = ConTblBKMRK.Offset(7, i).Value Then
        IOCcnt = 2
        ConTblBKMRK.Offset(7, i).Interior.Color = 16711935
    ElseIf ConTblBKMRK.Offset(7, i).Interior.Color = 16711935 Then
        ConTblBKMRK.Offset(7, i).Interior.ColorIndex = 0
    End If
Next



End Function
Function SpendReasonable() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
BenchBKMRK.End(xlDown).Offset(0, 1).Select

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    SpendReasonable = 1
Else
    SpendReasonable = 3
End If


End Function
Function PRS() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
prsBKMRK.Offset(MbrNMBR + 1, 0).Select

'<<Check>>
'=========================================
PRS = 1
For i = 1 To suppNMBR
    If ActiveCell.Offset(0, i * 2).Value > ActiveCell.Offset(0, i * 2 - 1).Value * 1.05 Or (ActiveCell.Offset(0, i * 2) = 0 And Not ActiveCell.Offset(0, i * 2 - 1).Value = 0) Then
        PRS = 2
        ActiveCell.Offset(1, i * 2).Interior.Color = 16711935
    ElseIf ActiveCell.Offset(1, i * 2).Interior.Color = 16711935 Then
        ActiveCell.Offset(1, i * 2).Interior.ColorIndex = 0
    End If
Next


End Function
Function AotrsLessThan() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
MSGraphBKMRK.End(xlDown).End(xlToRight).Offset(0, -1).Select

'<<Check>>
'=========================================
If ActiveCell.Value <= 0.05 Or Application.CountA(Range(MSGraphBKMRK.End(xlDown).Offset(0, 1), ActiveCell.Offset(0, -2))) / 2 >= 10 Then
    AotrsLessThan = 1
Else
    AotrsLessThan = 2
End If


End Function
Function MSgraphData() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
Range(MSGraphBKMRK.End(xlDown).Offset(0, 1), MSGraphBKMRK.End(xlDown).End(xlToRight)).Select

'<<Check>>
'=========================================
GrphStr = Sheets("Initiative Spend Overview").ChartObjects(2).Chart.SeriesCollection(1).Formula
GrphStr = Replace(GrphStr, "'Initiative Spend Overview'!", "")
GrphStr = Replace(GrphStr, "(", "")
GrphStr = Replace(GrphStr, ")", "")
GrphStr = Mid(GrphStr, 9, Len(GrphStr))
GrphStr = Left(GrphStr, Len(GrphStr) - 2)

If Round(WorksheetFunction.sum(Sheets("initiative spend overview").Range(GrphStr)), 1) = Round(MSGraphBKMRK.End(xlDown).End(xlToRight).Value, 1) Then
    MSgraphData = 1
Else
    MSgraphData = 2
End If


End Function
Function MSgraphReadable() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
MSGraphBKMRK.End(xlToRight).Select

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    MSgraphReadable = 1
Else
    MSgraphReadable = 3
End If


End Function
Function BenchGraphData() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
BenchBKMRK.End(xlDown).Select

'<<Check>>
'=========================================
For i = 1 To 3
    GrphStr = Sheets("Initiative Spend Overview").ChartObjects(1).Chart.SeriesCollection(i).Formula
    GrphStr = Replace(GrphStr, "'Initiative Spend Overview'!", "")
    GrphStr = Replace(GrphStr, "=SERIES", "")
    GrphStr = Replace(GrphStr, "(", "")
    GrphStr = Replace(GrphStr, ")", "")
    GrphStr = Replace(GrphStr, ",,", ",")
    GrphStr = Left(GrphStr, Len(GrphStr) - 2)
    ttlGrphStr = GrphStr & "," & ttlGrphStr
Next
ttlGrphStr = Left(ttlGrphStr, Len(ttlGrphStr) - 1)

If Round(WorksheetFunction.sum(Sheets("initiative spend overview").Range(ttlGrphStr)), 1) = Round(BenchBKMRK.End(xlDown).Offset(0, 3).Value + BenchBKMRK.End(xlDown).Offset(0, 5).Value + BenchBKMRK.End(xlDown).Offset(0, 7).Value, 1) Then
    BenchGraphData = 1
Else
    BenchGraphData = 2
End If


End Function
Function BenchGraphReadable() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Initiative Spend Overview")
BenchBKMRK.End(xlToRight).Select

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    BenchGraphReadable = 1
Else
    BenchGraphReadable = 3
End If


End Function
Function PropTotalsSum() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Vizient Contracts - Conv")
Range(ConvBKMRK.Offset((MbrNMBR + 1), 3), ConvBKMRK.Offset((MbrNMBR + 1), 6)).Select

'<<Check>>
'=========================================
PropTotalsSum = 1
For i = 1 To suppNMBR
    Set Propttl = Range(ConvBKMRK.Offset(MbrNMBR + 1 + (MbrNMBR + 8) * (i - 1), 4).Address & "," & ConvBKMRK.Offset(MbrNMBR + 1 + (MbrNMBR + 8) * (i - 1), 6).Address)
    If Not Round(WorksheetFunction.sum(Propttl), 1) = Round(ConvBKMRK.Offset(MbrNMBR + 1 + (MbrNMBR + 8) * (i - 1), 1).Value, 1) Then
        PropTotalsSum = 2
        Propttl.Interior.Color = 16711935
    ElseIf Propttl.Interior.Color = 16711935 Then
        ConvBKMRK.Offset(MbrNMBR + 1 + (MbrNMBR + 8) * (i - 1), 2).Copy
        Propttl.PasteSpecial xlPasteFormats
    End If
Next
Application.CutCopyMode = False


End Function
Function ConMftrMtchMPP() As Integer


Dim conn As New ADODB.Connection
Dim recset As New ADODB.Recordset
    
'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line item data")
Range("BG3").Select

'<<Check>>
'=========================================
    ConMftrMtchMPP = 1
    On Error GoTo ERR_NoDB

    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    
    For i = 0 To suppNMBR - 1
        conParam = conParam & "'" & UCase(ZeusForm.asscContracts.List(i)) & "',"
    Next
    conParam = Left(conParam, Len(conParam) - 1)

    SELECTstr = "SELECT DISTINCT OV.VENDOR_NAME, OC.CONTRACT_NUMBER "
    FROMstr = "FROM OCSDW_CONTRACT OC INNER JOIN OCSDW_VENDOR AS OV ON OC.VENDOR_KEY = OV.VENDOR_KEY  INNER JOIN OCSDW_PRICE_Novation AS OP ON OC.CONTRACT_ID = OP.CONTRACT_ID AND OP.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "' INNER JOIN OCSDW_PRICE_TIER AS OPT ON OP.PRICE_TIER_KEY = OPT.PRICE_TIER_KEY  INNER JOIN OCSDW_CONTRACT_PROGRAM_DETAIL cpd ON OC.CONTRACT_ID = cpd.CONTRACT_ID LEFT OUTER JOIN OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL cav ON cav.CONTRACT_ID = oc.CONTRACT_ID AND cav.attribute_value_id = '963' "
    WHEREstr = "WHERE OC.CONTRACT_NUMBER In (" & conParam & ") ORDER BY OC.CONTRACT_NUMBER ASC" 'AND (OC.Status_Key = 'ACTIVE' or OC.Status_Key = 'signed' or OC.Status_Key = 'expired')
    sqlstr = SELECTstr & FROMstr & WHEREstr

    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    recset.MoveFirst

    For i = 1 To suppNMBR
        If Not Range("BG3").Offset(0, 30 * (i - 1)).Value = recset.Fields(0) Then
            ConMftrMtchMPP = 2
            Range("BG3").Offset(0, 30 * (i - 1)).Interior.Color = 16711935
        ElseIf Range("BG3").Offset(0, 30 * (i - 1)).Interior.Color = 16711935 Then
            Range("BG3").Offset(0, 30 * (i - 1)).Interior.ColorIndex = 0
        End If
        recset.MoveNext
    Next
    
Exit Function
'::::::::::::::::::::::::::::::::::::;
ERR_NoDB:
    ConMftrMtchMPP = 3
    ManualCheck = True
    NoteUnhide = True
    NoteText = "Could not connect to Database, must be checked manually."
Exit Function



End Function
Function NonConMftrNmsStd() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line item data")
Range("V:V").Select

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    NonConMftrNmsStd = 1
Else
    NonConMftrNmsStd = 3
End If


End Function
Function OneMftrPerCatnum() As Integer

On Error Resume Next

'<<Check>>
'=========================================
QCFlg = True
Call StandardizeMfg  '>>>>>>>>>>
QCFlg = False

If QCChkFlg = True Then
    OneMftrPerCatnum = 1
Else
    OneMftrPerCatnum = 2
End If

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line item data")
Range("V:V").Select



End Function
Function BlankMftrs(PrevStatus As Boolean) As Integer


'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line item data")
Range("V:V").Select

'<<Checks>>
'=========================================

'Check Unknowns
'------------------
BlankMftrs = 1
On Error GoTo ERR_NoUkn
Selection.Find(what:="Unknown", lookat:=xlPart).Select
BlankMftrs = 3
If PrevStatus = True Then
    BlankMftrs = 1
Else
    NoteUnhide = True
    NoteText = "If Unknowns have been thoroughly checked and you still can't find manufacturer name then click again."
End If
On Error Resume Next

'check if blanks
'-------------------------
ChkBlnks:
If WorksheetFunction.CountIf(Range("V5:V" & Range("A4").End(xlDown).Row), "") > 0 Or WorksheetFunction.CountIf(Range("V5:V" & Range("A4").End(xlDown).Row), "0") > 0 Then BlankMftrs = 2


Exit Function
'::::::::::::::::::::::::::
ERR_NoUkn:
Resume ChkBlnks



End Function
Function CatnumsStdzd() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line item data")
Range("X:X").Select

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    CatnumsStdzd = 1
Else
    CatnumsStdzd = 3
End If


End Function
Function BlankCatnums() As Integer

On Error Resume Next

'<<Navigation>>
'=========================================
Call FUN_TestForSheet("Line item data")
Range("X:X").Select

'<<Check>>
'=========================================
If WorksheetFunction.CountIf(Range("X5:X" & Range("A4").End(xlDown).Row), "") > 0 Or WorksheetFunction.CountIf(Range("X5:X" & Range("A4").End(xlDown).Row), "0") > 0 Then
    BlankCatnums = 2
Else
    BlankCatnums = 1
End If


End Function
Function PlvlVars() As Integer

    On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Line item data")
    Range("AO:AO").Select
    
    '<<Check>>
    '=========================================
    Call Calculate_Priceleveling  '>>>>>>>>>>>
    plvlVarRange = Val(Format(ZeusForm.plvlRngSet.Value, "0.00"))

    PlvlVars = 1
    For Each c In Range("AO5:AO" & Range("A4").End(xlDown).Row)
        If ((c.Value > plvlVarRange Or c.Value < -plvlVarRange) And Not LCase(Range("AI" & c.Row).Value) = "x") Or (c.Value > 2 Or c.Value < -2) Then
            PlvlVars = 2
            Range("AI" & c.Row).Interior.Color = 16711935
        ElseIf Range("AI" & c.Row).Interior.Color = 16711935 Then
            Range("AI" & c.Row).Interior.ColorIndex = 0
        End If
    Next


End Function
Function MbrBenchVars() As Integer

On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Line item data")
    Range("AO:AO").Select
    
    '<<Check>>
    '=========================================
    BnchVarRange = Val(Format(ZeusForm.bnchRngSet.Value, "0.00"))

    MbrBenchVars = 1
    For Each c In Range("AV5:AV" & Range("A4").End(xlDown).Row)
        If Trim(c.Value) = "" Then
        ElseIf ((c.Value > BnchVarRange Or c.Value < -BnchVarRange) And Not LCase(Range("AI" & c.Row).Value) = "x") Or (c.Value > 3 Or c.Value < -3) Then
            MbrBenchVars = 2
            Range("AV" & c.Row).Interior.Color = 16711935
        ElseIf Range("AV" & c.Row).Interior.Color = 16711935 Then
            Range("AV" & c.Row).Interior.ColorIndex = 0
        End If
    Next


End Function
Function SuppVars() As Integer

    On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Line item data")
    Range("BN:BN").Select
    
    '<<Check>>
    '=========================================
    suppVarRange = Val(Format(ZeusForm.SuppRngSet.Value, "0.00"))

    SuppVars = 1
    For i = 1 To suppNMBR
        For Each c In Range("BN5:BN" & Range("A4").End(xlDown).Row).Offset(0, 30 * (i - 1))
            If Trim(c.Value) = "" Then
            ElseIf ((c.Value > suppVarRange Or c.Value < -suppVarRange) And Not LCase(Range("AI" & c.Row).Value) = "x") Or (c.Value > 2 Or c.Value < -2) Then
                SuppVars = 2
                Range("BN" & c.Row).Offset(0, 30 * (i - 1)).Interior.Color = 16711935
            ElseIf Range("BN" & c.Row).Offset(0, 30 * (i - 1)).Interior.Color = 16711935 Then
                Range("BN" & c.Row).Offset(0, 30 * (i - 1)).Interior.ColorIndex = 0
            End If
        Next
    Next


End Function
Function SuppBenchVars() As Integer

    On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Line item data")
    Range("BZ:BZ").Select
    
    '<<Check>>
    '=========================================
    BnchVarRange = Val(Format(ZeusForm.bnchRngSet.Value, "0.00"))

    SuppBenchVars = 1
    For i = 1 To suppNMBR
        For Each c In Range("BZ5:BZ" & Range("A4").End(xlDown).Row).Offset(0, 30 * (i - 1))
            If Trim(c.Value) = "" Then
            ElseIf ((c.Value > BnchVarRange Or c.Value < -BnchVarRange) And Not LCase(Range("AI" & c.Row).Value) = "x") Or (c.Value > 3 Or c.Value < -3) Then
                SuppBenchVars = 2
                Range("BZ" & c.Row).Offset(0, 30 * (i - 1)).Interior.Color = 16711935
            ElseIf Range("BZ" & c.Row).Offset(0, 30 * (i - 1)).Interior.Color = 16711935 Then
                Range("BZ" & c.Row).Offset(0, 30 * (i - 1)).Interior.ColorIndex = 0
            End If
        Next
    Next


End Function
Function PkgMsmtch() As Integer


    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Line item data")
    Range("AC:AC").Select
    
    '<<Check>>
    '=========================================
    PkgMsmtch = 1
    MsmtchVar = Val(Format(ZeusForm.msmtchSelect.Value, "0.00"))
    Set AlreadyChkd = Range("A4").End(xlToRight).Offset(0, 1)
    ItmNmbr = Application.CountA(Range("X:X")) - 1
    For i = 1 To ItmNmbr + 1
        If Not AlreadyChkd.Offset(i, 0).Value = 1 Then
            catnum = Range("X4").Offset(i, 0).Value
            mtchcats = Application.CountIf(Range("X:X"), catnum)
                
            'find itms assc with catnum
            '---------------------------------
            EAcnt = 0
            ReDim CatArray(1 To 1)
            Set srchstrt = Range("X4").Offset(i - 1, 0)
            On Error GoTo ERR_NxtItm
            For j = 1 To mtchcats
                Set srchstrt = Range(srchstrt, srchstrt.End(xlDown)).Find(what:=catnum, lookat:=xlWhole)
                ReDim Preserve CatArray(1 To j)
                Set CatArray(j) = srchstrt
                AlreadyChkd.Offset(srchstrt.Row - 4).Value = 1
                If srchstrt.Offset(0, 5).Value = "EA" Then EAcnt = EAcnt + 1
NxtItm:
            Next
            On Error GoTo 0
            
            'Evaluate pkg mismatches
            '---------------------------------
            For j = 1 To mtchcats
                Set itm = CatArray(j)
                If itm.Offset(0, 4).Value = 1 Then
                    If itm.Offset(0, 5).Value = "BN" Or itm.Offset(0, 5).Value = "CA" Or itm.Offset(0, 5).Value = "PL" Or itm.Offset(0, 5).Value = "BG" Or itm.Offset(0, 5).Value = "BX" Or itm.Offset(0, 5).Value = "CT" Or itm.Offset(0, 5).Value = "DZ" Or itm.Offset(0, 5).Value = "PK" Then
                        If EAcnt / mtchcats < MsmtchVar Then
                            PkgMsmtch = 3
                            Range("AC" & itm.Row).Interior.Color = 16711935
                        ElseIf Range("AC" & itm.Row).Interior.Color = 16711935 Then
                            Range("AC" & itm.Row).Interior.ColorIndex = 0
                        End If
                    End If
                ElseIf Range("AC" & itm.Row).Interior.Color = 16711935 Then
                    Range("AC" & itm.Row).Interior.ColorIndex = 0
                End If
            Next
        
        End If
    Next
    
    Range(AlreadyChkd, AlreadyChkd.Offset(ItmNmbr, 0)).ClearContents

    If PkgMsmtch = 3 Then
        NoteUnhide = True
        NoteText = "Even if a UOM quantitiy isn't creating a significant variance, it's supposed to match up with the associated pkg description.  (Ex. If a pkg description in column AC is ""EA"" for each, then the UOM quantity should be 1.  Likewise if a pkg description is CA then the UOM quantity should probably be greater than 1) At the same time be aware that, although statistically less often, pkg descriptions can be wrong as well.  So if you've researched a UOM quantity and you strongly believe the UOM quantity for an item to be correct but the pkg description doesn't seem to match then that's fine, the point of the pkg description is to steer you in the right direction of the UOM quantitiy.  If that's the case though you might want to note it on the notes tab to CYA."
    End If

Exit Function
':::::::::::::::::::::::::::::::::;;
ERR_NxtItm:
Resume NxtItm



End Function
Function OOSRmvd(PrevStatus As Boolean) As Integer

    On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Items Removed")
    Range("A1").Select

    '<<Check>>
    '=========================================
    If Not Application.CountA(Range("A:A")) - 1 = 0 Then
        OOSRmvd = 1
    Else
        OOSRmvd = 3
        If PrevStatus = True Then
            OOSRmvd = 1
        Else
            NoteUnhide = True
            NoteText = "If there's really no items that need to be removed then click again."
        End If
    End If
    

End Function
Function Novaplus() As Integer

    On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Index")
    Range(ConTblBKMRK.Offset(6, 1), ConTblBKMRK.Offset(6, suppNMBR)).Select

    '<<Check>>
    '=========================================
    Novaplus = 1
    For i = 1 To suppNMBR
        If ConTblBKMRK.Offset(6, i) = "ü" Then
            For Each c In Range(LIDSuppBKMRK.Offset(1, 30 * (i - 1)), LIDSuppBKMRK.Offset(0, 30 * (i - 1)).End(xlDown))
                If c.HasFormula And InStr(c.Value, "V") = 1 Then
                    Novaplus = 2
                    c.Interior.Color = 16711935
                ElseIf c.Interior.Color = 16711935 Then
                    c.Interior.ColorIndex = 0
                End If
            Next
        End If
    Next


End Function
Function DupCat() As Integer


    '<<Check>>
    '=========================================
    DupCat = 1
    For i = 1 To suppNMBR
        Call FUN_TestForSheet(FUN_SuppName(i) & " Pricing")
        If Not Application.CountA(Range("A:A")) - 1 = 0 Then
            For Each c In Range("A2:A" & Range("F1").End(xlDown).Row)
                If WorksheetFunction.CountIf(Range("A:A"), c.Value) > 1 Then
                    DupCat = 2
                    c.Interior.Color = 16711935
                ElseIf c.Interior.Color = 16711935 Then
                    c.Interior.ColorIndex = 0
                End If
            Next
        End If
    Next
    
    If DupCat = 2 Then
        NoteUnhide = True
        NoteText = "If there are duplicate catalog numbers the vlookup formula used to pull in pricing on the line item data tab will only pull in pricing for the first item.  Please make sure this is what you want or remove all but one item per catalog number."
    End If
    

End Function
Function BestPriceFrmla() As Integer


    '<<Check>>
    '=========================================
    BestPriceFrmla = 1
    For i = 1 To suppNMBR
        Call FUN_TestForSheet(FUN_SuppName(i) & " Pricing")
        If Not Application.CountA(Range("A:A")) - 1 = 0 Then
            TierNmbr = Rows("1:1").Find(what:="EA Price").Column - Range("J1").Column - 1
            If ConTblBKMRK.Offset(8, i).Value = "Best Price" Then
                For Each c In Range("J2:J" & Range("F1").End(xlDown).Row)
                    If Not c.Value = WorksheetFunction.Min(Range(c, c.Offset(0, TierNmbr))) Then
                        BestPriceFrmla = 2
                        c.Interior.Color = 16711935
                    ElseIf c.Interior.Color = 16711935 Then
                        c.Interior.ColorIndex = 0
                    End If
                Next
            ElseIf InStr(ConTblBKMRK.Offset(8, i).Value, "Tier") > 0 Then
                TierOffset = Int(Replace(ConTblBKMRK.Offset(8, i).Value, "Tier ", ""))
                For Each c In Range("J2:J" & Range("F1").End(xlDown).Row)
                    If Not c.Value = c.Offset(0, TierOffset) Then
                        BestPriceFrmla = 2
                        c.Interior.Color = 16711935
                    ElseIf c.Interior.Color = 16711935 Then
                        c.Interior.ColorIndex = 0
                    End If
                Next
            End If
        End If
    Next


End Function
Function ZeroItmsRmvd() As Integer


    '<<Check>>
    '=========================================
    ZeroItmsRmvd = 1
    For i = 1 To suppNMBR
        Call FUN_TestForSheet(FUN_SuppName(i) & " Pricing")
        If Not Application.CountA(Range("A:A")) - 1 = 0 Then
            Range("J:J").Calculate
            For Each c In Range("J2:J" & Range("F1").End(xlDown).Row)
                If c.Value = 0 Or Trim(c.Value) = "" Then
                    ZeroItmsRmvd = 2
                    c.Interior.Color = 16711935
                ElseIf c.Interior.Color = 16711935 Then
                    c.Interior.ColorIndex = 0
                End If
            Next
        End If
    Next



End Function
Function UnqualTrsRmvd() As Integer

    
    '<<Check>>
    '=========================================
    UnqualTrsRmvd = 1
    For i = 1 To suppNMBR
        Call FUN_TestForSheet(FUN_SuppName(i) & " Pricing")
        If Not Application.CountA(Range("A:A")) - 1 = 0 Then
            TierNmbr = Rows("1:1").Find(what:="EA Price").Column - Range("J1").Column - 1
            If Not Application.CountA(Range(ConTblBKMRK.Offset(9, i), ConTblBKMRK.End(xlDown).Offset(0, i))) = TierNmbr Then UnqualTrsRmvd = 2
        End If
    Next


End Function
Function CurrPricingUsed() As Integer

    On Error Resume Next

    '<<Navigation>>
    '=========================================
    Call FUN_TestForSheet("Index")
    Range(ConTblBKMRK.Offset(4, 1), ConTblBKMRK.Offset(4, suppNMBR)).Select

    '<<Check>>
    '=========================================
    CurrPricingUsed = 1
    For Each c In Selection
        If Date > c.Value Then
            CurrPricingUsed = 2
            c.Interior.Color = 16711935
        ElseIf c.Interior.Color = 16711935 Then
            c.Interior.ColorIndex = 0
        End If
    Next


End Function
Function XrefSorted() As Integer


    '<<Check>>
    '=========================================
    XrefSorted = 1
    For i = 1 To suppNMBR
        Call FUN_TestForSheet(FUN_SuppName(i) & " Cross Reference")
        If Not Trim(Range("F2").Value) = "" Then
            
            On Error GoTo EndClean
            Set FirstCat = Range("A1")
            mtchcats = 1
            Do
                Set FirstCat = FirstCat.Offset(mtchcats, 0)
                If Trim(FirstCat.Value) = "" Then Set FirstCat = FirstCat.End(xlDown)
                mtchcats = WorksheetFunction.CountIf(Range("A:A"), FirstCat.Value)
                Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))
                If Not mtchcats = 1 Then
                
                    For Each c In AllCats
                        If Not c.Value = FirstCat.Value Then
                            XrefSorted = 2
                            GoTo EndClean
                        ElseIf Not c.Row = FirstCat.Offset(mtchcats - 1, 0).Row Then
                            If InStr(LCase(c.Offset(0, 5).Value), "intelli") > 0 And InStr(LCase(c.Offset(1, 5).Value), "core") > 0 Then
                                XrefSorted = 2
                                GoTo EndClean
                            End If
                            If Not IsError(c.Offset(0, 8).Value) And Not IsError(c.Offset(1, 8).Value) Then
                                If c.Offset(0, 8).Value > c.Offset(1, 8).Value Then
                                    XrefSorted = 2
                                    GoTo EndClean
                                End If
                            End If
                        End If
                    Next
                    
                End If
            Loop Until FirstCat.Offset(mtchcats, 0).Row > Range("F1").End(xlDown).Row
        
        End If
    Next


Exit Function:
':::::::::::::::::::::::::::::::::::::
EndClean:
    NoteUnhide = True
    NoteText = "Supplier Cross reference tabs should be sorted 1st by Member Catalog Number ascending, then by Source ascending, then by EA price ascending."



End Function
Function XrefCleansed() As Integer

On Error Resume Next

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    XrefCleansed = 1
Else
    XrefCleansed = 3
End If


End Function
Function AdminFees() As Integer

On Error Resume Next

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    AdminFees = 1
Else
    AdminFees = 3
End If


End Function
Function BMPSorted() As Integer

On Error Resume Next

    '<<Check>>
    '=========================================
    BMPSorted = 1
    Call FUN_TestForSheet("best market price")

    Set FirstCat = Range("A1")
    mtchcats = 1
    Do
        Set FirstCat = FirstCat.Offset(mtchcats, 0)
        If Trim(FirstCat.Value) = "" Then Set FirstCat = FirstCat.End(xlDown)
        mtchcats = WorksheetFunction.CountIf(Range("A:A"), FirstCat.Value)
        Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))
        If Not mtchcats = 1 Then
            '[sort by psc first then by part num]
'            For Each c In AllCats
'                If Not c.Value = FirstCat.Value Then
'                    BMPSorted = 2
'                    GoTo endclean
'                ElseIf Not c.Row = FirstCat.Offset(mtchcats - 1, 0).Row Then
'                    If c.Offset(0, 5).Value < c.Offset(1, 5).Value Then
'                        BMPSorted = 2
'                        GoTo endclean
'                    End If
'                End If
'            Next
            
        End If
    Loop Until FirstCat.Offset(mtchcats, 0).Row > Range("F1").End(xlDown).Row


Exit Function:
':::::::::::::::::::::::::::::::::::::
EndClean:
    NoteUnhide = True
    NoteText = "The data in the best market price tab should be sorted first by ""Part Number"" ascending, then by ""Sample Size"" descending."



End Function
Function FullyCalc() As Integer

On Error Resume Next

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    FullyCalc = 1
Else
    FullyCalc = 3
End If


End Function
Function FontsFormatted() As Integer


    FontsFormatted = 1

    'check index & initiative spend overview tables
    '-------------------
    FontsFormatted = ChkFonts(Range(MbrBkmrk, MbrBkmrk.End(xlDown)))
    FontsFormatted = ChkFonts(Range(MSGraphBKMRK, MSGraphBKMRK.End(xlToRight).End(xlDown)))
    FontsFormatted = ChkFonts(Range(BenchBKMRK, BenchBKMRK.End(xlToRight).End(xlDown)))
    FontsFormatted = ChkFonts(Range(prsBKMRK, prsBKMRK.End(xlToRight).End(xlDown)))
        
    'check supplier tables
    '-------------------------
    For i = 1 To suppNMBR
        FontsFormatted = ChkFonts(Range(NonConBKMRK.Offset((MbrNMBR + 8) * (i - 1), 0), NonConBKMRK.End(xlToRight).Offset((MbrNMBR + 8) * (i - 1), 0)))
        FontsFormatted = ChkFonts(Range(ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1), 0), ConvBKMRK.End(xlToRight).Offset((MbrNMBR + 8) * (i - 1), 0)))
    Next

    'check Line item data
    '-------------------
    FontsFormatted = ChkFonts(Range(Sheets("Line item data").Range("A4"), Sheets("Line item data").Range("A4").End(xlDown).Offset(0, 57 + 30 * suppNMBR)))



End Function
Function OrgnlsPosted() As Integer

On Error Resume Next

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    OrgnlsPosted = 1
Else
    OrgnlsPosted = 3
End If


End Function
Function DataIntact() As Integer

On Error Resume Next

'<<Check>>
'=========================================
ManualCheck = True
NoteUnhide = True
NoteText = "No automated check, must check manually"

If Not QCMain = True Then
    DataIntact = 1
Else
    DataIntact = 3
End If


End Function
Function NoHardcoded() As Integer

    NoHardcoded = 1

    'check index & initiative spend overview tables
    '-------------------
    If ChkHardcode(Range(MSGraphBKMRK.Offset(1, -1), MSGraphBKMRK.End(xlToRight).Offset(MbrNMBR, 0))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(MSGraphBKMRK.Offset(MbrNMBR + 1, 1), MSGraphBKMRK.End(xlToRight).Offset(MbrNMBR + 1, 0))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(BenchBKMRK.Offset(1, -1), BenchBKMRK.Offset(MbrNMBR, 1))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(BenchBKMRK.Offset(1, 3), BenchBKMRK.End(xlToRight).Offset(MbrNMBR + 1, 0))) = 2 Then NoHardcoded = 2
    
    On Error Resume Next
    For Each c In Range(prsBKMRK.Offset(0, -1), prsBKMRK.End(xlToRight))
        If Not InStr(c.Value, "Reported") > 0 Then
            If ChkHardcode(Range(c.Offset(1, 0), c.Offset(MbrNMBR, 0))) = 2 Then NoHardcoded = 2
        End If
    Next
    On Error GoTo 0
    
    'check supplier tables
    '-------------------------
    For i = 1 To suppNMBR
        If ChkHardcode(Range(NonConBKMRK.Offset((MbrNMBR + 8) * (i - 1) + 1, -1), NonConBKMRK.End(xlToRight).Offset((MbrNMBR + 8) * (i - 1) + MbrNMBR, 0))) = 2 Then NoHardcoded = 2
        If ChkHardcode(Range(NonConBKMRK.Offset((MbrNMBR + 8) * (i - 1) + MbrNMBR + 1, 1), NonConBKMRK.End(xlToRight).Offset((MbrNMBR + 8) * (i - 1) + MbrNMBR + 1, 0))) = 2 Then NoHardcoded = 2
        If ChkHardcode(Range(ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1) + 1, -1), ConvBKMRK.End(xlToRight).Offset((MbrNMBR + 8) * (i - 1) + MbrNMBR, 0))) = 2 Then NoHardcoded = 2
        If ChkHardcode(Range(ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1) + MbrNMBR + 1, 1), ConvBKMRK.End(xlToRight).Offset((MbrNMBR + 8) * (i - 1) + MbrNMBR + 1, 0))) = 2 Then NoHardcoded = 2
    Next

    'check Line item data
    '-------------------
    If ChkHardcode(Range(Sheets("Line item data").Range("Z5"), Sheets("Line item data").Range("Z" & Sheets("Line item data").Range("A4").End(xlDown).Row))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(Sheets("Line item data").Range("AG5"), Sheets("Line item data").Range("AH" & Sheets("Line item data").Range("A4").End(xlDown).Row))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(Sheets("Line item data").Range("AJ5"), Sheets("Line item data").Range("AJ" & Sheets("Line item data").Range("A4").End(xlDown).Row))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(Sheets("Line item data").Range("AL5"), Sheets("Line item data").Range("AL" & Sheets("Line item data").Range("A4").End(xlDown).Row))) = 2 Then NoHardcoded = 2
    If ChkHardcode(Range(Sheets("Line item data").Range("AN5"), Sheets("Line item data").Range("BF" & Sheets("Line item data").Range("A4").End(xlDown).Row))) = 2 Then NoHardcoded = 2
    For i = 1 To suppNMBR
        If ChkHardcode(Range(Sheets("Line item data").Range("BG5"), Sheets("Line item data").Range("BQ" & Sheets("Line item data").Range("A4").End(xlDown).Row).Offset(0, 30 * (suppNMBR - 1)))) = 2 Then NoHardcoded = 2
        If ChkHardcode(Range(Sheets("Line item data").Range("BU5"), Sheets("Line item data").Range("CJ" & Sheets("Line item data").Range("A4").End(xlDown).Row).Offset(0, 30 * (suppNMBR - 1)))) = 2 Then NoHardcoded = 2
    Next


End Function


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SECONDARY FUNCTIONS ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function ChkFonts(ChkRng As Range)
    
    ChkFonts = 1
    For Each c In ChkRng
        If Not LCase(c.Font.Name) = "arial" Or Not c.Font.Size = 8 Then
            ChkFonts = 2
            c.Interior.Color = 16711935
        ElseIf c.Interior.Color = 16711935 Then
            c.Interior.ColorIndex = 0
        End If
    Next
    
End Function
Function ChkHardcode(ChkRng As Range)
    
    ChkHardcode = 1
    For Each c In ChkRng
        If Not c.HasFormula Then
            ChkHarcode = 2
            c.Interior.Color = 16711935
        ElseIf c.Interior.Color = 16711935 Then
            c.Interior.ColorIndex = 0
        End If
    Next
    
End Function
Function Check_ttls(Chk As String)


On Error Resume Next

MSgraph_ttl = Round(MSGraphBKMRK.End(xlDown).End(xlToRight).Value, 0)
If MSgraph_ttl = "" Then MSgraph_ttl = 0
Bench_ttl = Round(BenchBKMRK.End(xlDown).Offset(0, 1).Value, 0)
If Bench_ttl = "" Then Bench_ttl = 0
NC_ttl = Round(NonConBKMRK.End(xlDown).Offset(-1, 1).Value, 0)
If NC_ttl = "" Then NC_ttl = 0
Conv_ttl = Round(ConvBKMRK.End(xlDown).Offset(-1, 1).Value, 0)
If Conv_ttl = "" Then Conv_ttl = 0
LID_ttl = Round(Sheets("line item data").Range("AJ3").Value, 0)
If LID_ttl = "" Then LID_ttl = 0

For Each c In Range(prsBKMRK.Offset(0, 1), prsBKMRK.End(xlToRight))
    If Not InStr(c.Value, "Reported") > 0 Then PRS_ttl = PRS_ttl + Round(c.Offset(MbrNMBR + 1, 0).Value, 1)
Next
If PRS_ttl = "" Then PRS_ttl = 0
PRS_ttl = Round(PRS_ttl, 0)

Sheets("index").Range("B1").Formula = "=Mode(" & MSgraph_ttl & "," & PRS_ttl & "," & Bench_ttl & "," & NC_ttl & "," & Conv_ttl & "," & LID_ttl & ")"
Mode_ttl = Sheets("index").Range("B1").Value
Sheets("index").Range("B1").ClearContents
'Mode_ttl = WorksheetFunction.Mode(MSgraph_ttl, PRS_ttl, Bench_ttl, NC_ttl, Conv_ttl, LID_ttl)

'ChkStatus:
If Chk = "MSgraph" Then
    If MSgraph_ttl = Mode_ttl Then Check_ttls = True
ElseIf Chk = "Bench" Then
    If Bench_ttl = Mode_ttl Then Check_ttls = True
ElseIf Chk = "PRS" Then
    If PRS_ttl = Mode_ttl Then Check_ttls = True
ElseIf Chk = "NC" Then
    If NC_ttl = Mode_ttl Then Check_ttls = True
ElseIf Chk = "Conv" Then
    If Conv_ttl = Mode_ttl Then Check_ttls = True
ElseIf Chk = "LID" Then
    If LID_ttl = Mode_ttl Then Check_ttls = True
End If

On Error GoTo 0
Exit Function
':::::::::::::::::::::::::
'ERR_ttl:
'Resume ChkStatus


End Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'OTHER QC BUTTONS AND SUBS ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub unlockQC(LockVal As Boolean) '(control As IRibbonControl)


If LockVal = True Then
    If LCase(Usr) = "bforrest" Then
        ActiveSheet.Unprotect Password = "existentialism"
    ElseIf LCase(Usr) = "rzhang" Or LCase(Usr) = "dlovelace" Then
        If ReviewFlg = 1 Then
            ActiveSheet.Unprotect Password = "existentialism"
        Else
            MsgBox "This feature is only available if you are acting as a reviewer"
            Exit Sub
        End If
    Else
        MsgBox "You do not have permission to unlock QC."
        Exit Sub
    End If
Else
    ActiveSheet.Protect Password = "existentialism"
End If


End Sub
Sub QCHelp(HelpVal As Boolean)

On Error Resume Next

    Sheets("QC").Visible = True
    Sheets("QC").Select
    Sheets("QC").Unprotect Password = "existentialism"
    
    If HelpVal = True Then
        Sheets("QC").Range("A:B").EntireColumn.Hidden = False
        Range("A1").Select
    Else
        Sheets("QC").Range("A:B").EntireColumn.Hidden = True
    End If
    
    Sheets("QC").Protect Password = "existentialism"


End Sub
Sub QCReview()


'Save to Reveiw folders
'=============================================================================================================================================
If InStr(LCase(ActiveWorkbook.Name), "(post qc)") > 0 Or InStr(LCase(ActiveWorkbook.Name), "(final)") > 0 Then GoTo NoVPN
    
    Application.DisplayAlerts = False
    dltNm = ActiveWorkbook.Name
    svnm = Replace(ActiveWorkbook.Name, ".xlsx", "")
    svnm = Replace(svnm, "(POST QC)", "")
    svnm = Replace(svnm, "(Initial)", "")
    svnm = Replace(svnm, "(PreQC)", "")

    On Error GoTo errhndlNOVPN
    CompletePATH = Replace(TMreviewPATH, "Ready for review", "Completed Review")
    ReadyPATH = TMreviewPATH
    
    'save to completed folder
    '---------------------------
    If Dir(CompletePATH & svnm & "(POST QC).xlsx", vbDirectory) = vbNullString Then ActiveWorkbook.SaveAs Filename:=CompletePATH & svnm & "(POST QC).xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    'copy to archive
    '---------------------------
    If Dir(CompletePATH & "Archive\" & svnm & "(POST QC).xlsx", vbDirectory) = vbNullString Then FileCopy ReadyPATH & dltNm, CompletePATH & "Archive\" & svnm & "(POST QC).xlsx"
        
    'Delete from ready folder
    '---------------------------
    If Not Dir(ReadyPATH & dltNm, vbDirectory) = vbNullString Then Kill ReadyPATH & dltNm

NoVPN:
    On Error GoTo 0
    Application.DisplayAlerts = True

'run QC checklist
'===============================================================================================================================
tmWB.Activate
QCform.Show (False)
QCform.ScrollTop = 0
Call QC_Main
Sheets("QC").Range("K5").Value = Usr


Exit Sub:
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNOVPN:
Resume NoVPN

End Sub
Sub QCsubmit()

'save a copy to Zeus,format, and clinical
'----------------------------------------------
Application.DisplayAlerts = False
svnm = Replace(ActiveWorkbook.Name, ".xlsx", "")
svnm = Replace(ActiveWorkbook.Name, "(PreQC)", "")
svnm = Replace(ActiveWorkbook.Name, "(Initial)", "")
ChDir ZeusPATH
ActiveWorkbook.SaveAs Filename:=ZeusPATH & "\" & svnm & "(PreQC).xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.Close (False)
FileCopy ZeusPATH & "\" & svnm & "(PreQC).xlsx", TMreviewPATH & "\" & svnm & "(PreQC).xlsx"


End Sub
Sub PostToIMETH()

'Check sort if marked as reviewer
'=======================================================================================================================================================
If ReviewFlg = 1 Then
    
    '[TBD]connect to initial extract and asf extract via adodb connection
    'check member names in asf extract same as you do in inititiative extract
    
    'Find extract wb
    '-------------------------------
    Call SetNetPATH(NetNm)
    On Error GoTo errhndlFndExWB
    dteFile = DateSerial(1900, 1, 1)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data\")

    extractNm = "initiative_extract"
    For Each ofile In objFolder.Files
        If InStr(LCase(ofile.Name), extractNm) > 0 And ofile.DateCreated > dteFile And Not InStr(ofile.Name, "$") > 0 Then
            dteFile = ofile.DateCreated
            ASFextractSvNm = ofile.Name
        End If
    Next
    If ASFextractSvNm = vbNullString Then
        MsgBox "Please select Initial Extract."
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        intchoice = Application.FileDialog(msoFileDialogOpen).Show
        If intchoice <> 0 Then
            wbstr = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
            Set exwb = Workbooks.Open(wbstr)
        Else
            GoTo RFsetup
        End If
    Else
        Set exwb = Workbooks.Open(NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data\" & ASFextractSvNm)
    End If
        
    'check Member, description, and unit price
    '------------------------------------------
FndExWB:
    If wbfoundFLG = 0 Then GoTo RFsetup
    On Error GoTo errhndlManualSpot
    tmWB.Activate
    For Each sht In tmWB.Sheets
        If LCase(sht.Name) = "sna standardization" Then
            sht.Visible = True
            sht.Select
            SNAstdFND = 1
            Exit For
        End If
    Next
    If Not SNAstdFND = 1 Then
        tmWB.Activate
        Call Import_StdznIndex
    End If
    
    lstName = Range("A1000").End(xlUp).Row
    lstStudy = Range("C1000").End(xlUp).Row
    Set wfnamerng = Range(Range("A4"), Range("A" & lstName))
    Set wfstudyrng = Range(Range("C4"), Range("C" & lstStudy))
    Set fndrng = Range(exwb.Sheets(1).Range("A2"), exwb.Sheets(1).Range("A1").End(xlDown))
    ActiveSheet.Visible = False
    Sheets("Line Item Data").Select
    Set ExtChkRng = Range(Range("A3"), Range("A" & Range("N2").End(xlDown).Row))
    On Error GoTo 0
    For Each c In ExtChkRng
        If Not Range("N" & c.Row).Value = 0 And Not Range("R" & c.Row).Value = 0 Then
            Set orgnlID = fndrng.Find(what:=Trim(c.Value), lookat:=xlWhole)
            If Not IsError(orgnlID.Offset(0, 5).Value) And Not Trim(Range("F" & c.Row).Value) = "" Then
                'Check hospital
                '-------------------
                On Error GoTo errhndlTRYNAME
                Resume
                StdName = Trim(LCase(wfstudyrng.Find(what:=orgnlID.Offset(0, 59).Value, lookat:=xlWhole).Offset(0, -1).Value))
stdMbr:         On Error GoTo errhndlManualSpot
                If Trim(LCase(Range("F" & c.Row).Value)) = StdName Then           '<<member (Can't use bc standardization of names)
                    'Check description
                    '-------------------
                    If Trim(LCase(Range("H" & c.Row).Value)) = Trim(LCase(fndrng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, 7).Value)) Then             '<<description (Can't use bc of rollup)
                        'Check unit cost
                        '-------------------
                        If Not Format(Trim(LCase(Range("R" & c.Row).Value)), "$0.00") = Format(Trim(LCase(fndrng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, 16).Value)), "$0.00") Then  '<<Unit price
                            extflg = extflg + 1
                            c.Interior.Color = 16711935
                        End If
                    Else
                        extflg = extflg + 1
                        c.Interior.Color = 16711935
                    End If
                Else
                    extflg = extflg + 1
                    c.Interior.Color = 16711935
                End If
            Else
                extflg = extflg + 1
                c.Interior.Color = 16711935
            End If
        End If
    Next
    
    'analyze results and mark on QC tab
    '---------------------------
    For Each sht In tmWB.Sheets
        If InStr(UCase(sht.Name), "QC") Then
            
            QCshtfnd = 1
            On Error GoTo errhndlManualSpot
            sht.Unprotect Password = "existentialism"
            
            If extflg > 0 Then
                  
                'check if > 2.5%
                '----------------------------
                If extflg > ExtChkRng.Count * 0.025 Then
                    sht.Range("C:C").Find(what:="matched to Extract", lookat:=xlPart).Offset(0, 1).Interior.Color = 65535
                    Exit For
                Else
                
                    'check to see if 3 highlights in a row
                    '----------------------------
                    For Each oID In ExtChkRng
                        If oID.Interior.Color = 16711935 Then
                            If oID.Offset(1, 0).Interior.Color = 16711935 Then
                                If oID.Offset(2, 0).Interior.Color = 16711935 Then
                                    sht.Range("C:C").Find(what:="matched to Extract", lookat:=xlPart).Offset(0, 1).Interior.Color = 65535
                                    GoTo RFsetup
                                End If
                            End If
                        End If
                    Next
                    sht.Range("C:C").Find(what:="matched to Extract", lookat:=xlPart).Offset(0, 1).Interior.Color = 65280
                    exwb.Close (False)
                    Exit For
                End If
            Else
                sht.Range("C:C").Find(what:="matched to Extract", lookat:=xlPart).Offset(0, 1).Interior.Color = 65280
                exwb.Close (False)
                Exit For
            End If
        End If
    Next
    If Not QCshtfnd = 1 Then GoTo QCend
    
RFsetup:
    '---------------------------
    On Error GoTo 0
    tmWB.Activate
    
Else

'Post/Check for report files
'=======================================================================================================================================================
'=======================================================================================================================================================

    'find file name for each componenet
    '------------------------------------
    tmWB.Activate
    Sheets("Line Item Data").Select
    dteFile = DateSerial(1900, 1, 1)
    Call setobjFolder(ZeusPATH)
    For Each ofile In objFolder.Files
        If InStr(ofile.Name, "CoreXref") > 0 And InStr(LCase(ofile.Name), LCase(FileName_PSC)) > 0 And Not InStr(ofile.Name, "$") > 0 Then
            CoreXrefSvNm = ofile.Name
        ElseIf InStr(LCase(ofile.Name), "dirt asf extract") > 0 And InStr(LCase(ofile.Name), LCase(FileName_PSC)) > 0 And Not InStr(ofile.Name, "$") > 0 Then
            ASFextractSvNm = ofile.Name
        ElseIf ofile.DateCreated > dteFile And InStr(LCase(ofile.Name), LCase(Usr)) > 0 And InStr(ofile.Name, "Initiative_Extract") > 0 And Not InStr(ofile.Name, "$") > 0 Then
            dteFile = ofile.DateCreated
            RawExtractSvNm = ofile.Name
        End If
    Next
    Call SetNetPATH(NetNm) '>>>>>>>>>>

End If

'FileCheckStrt:
'-----------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Application.DisplayAlerts = False

    'if no initiative folder then make one
    '-----------------------------------
    If Dir(NetworkPath & FileName_PSC, vbDirectory) = vbNullString Then
        If NetNm = "Northeast PC" Then NetworkPath = Replace(NetworkPath, "\Initiatives", "")
        MkDir (NetworkPath & FileName_PSC)
        MkDir (NetworkPath & FileName_PSC & "\1 DIR Files")
        MkDir (NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data")
        MkDir (NetworkPath & FileName_PSC & "\1 DIR Files\3 RFE & X-ref Files\")
        
        'MsgBox "No initiative folder found for this PSC.  Please submit your report manually."
        'retval = Shell("explorer.exe " & NetworkPath, vbNormalFocus)
        'WinOpenFLG = 1
    End If
    
    'Find initiative folder in network folder
    '-----------------------------------
    If Dir(NetworkPath & FileName_PSC & "\1 DIR Files", vbDirectory) = vbNullString Then
        MsgBox "No DIR folder found for this PSC.  Please submit your report manually."
        retval = Shell("explorer.exe " & NetworkPath & FileName_PSC, vbNormalFocus)
        WinOpenFLG = 1
    Else

        'extract
        '----------------
        If Dir(NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data\", vbDirectory) = vbNullString Then
            MsgBox "No HCO spend folder found for this PSC.  Please save your extract to the initiative folder manually."
            retval = Shell("explorer.exe " & NetworkPath & FileName_PSC & "\1 DIR Files\", vbNormalFocus)
            WinOpenFLG = 1
        Else
            If Not ReviewFlg = 1 Then
                On Error Resume Next
                If Not IsEmpty(RawExtractSvNm) Then FileCopy ZeusPATH & RawExtractSvNm, NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data\" & RawExtractSvNm
                If Not IsEmpty(ASFextractSvNm) Then FileCopy ZeusPATH & ASFextractSvNm, NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data\" & ASFextractSvNm
                On Error GoTo 0
            Else
                Set objFolder = objFSO.GetFolder(NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data")
                For Each ofile In objFolder.Files
                    If InStr(LCase(ofile.Name), "initiative_extract") Then
                        initexflg = 1
                    ElseIf InStr(LCase(ofile.Name), "dirt asf extract") Then
                        asfflg = 1
                    End If
                Next
                If Not initexflg = 1 Or Not asfflg = 1 Then
                    For Each sht In tmWB.Sheets
                        If InStr(UCase(sht.Name), "QC") Then
                            sht.Visible = True
                            sht.Select
                            sht.Unprotect Password = "existentialism"
                            Range("C:C").Find(what:="files posted", lookat:=xlPart).Offset(0, 1).Interior.Color = 65535
                            retval = Shell("explorer.exe " & NetworkPath & FileName_PSC & "\1 DIR Files\1 HCO Spend Data", vbNormalFocus)
                            GoTo QCend
                        End If
                    Next
                    GoTo QCend
                End If
            End If
        End If
        
        'xref
        '----------------
        If Dir(NetworkPath & FileName_PSC & "\1 DIR Files\3 RFE & X-ref Files\", vbDirectory) = vbNullString Then
            If Not WinOpenFLG = 1 Then
                MsgBox "No xref folder found for this PSC.  Please save your xref to the initiative folder manually."
                retval = Shell("explorer.exe " & NetworkPath & FileName_PSC & "\1 DIR Files\", vbNormalFocus)
            End If
        Else
            If Not ReviewFlg = 1 Then
                On Error Resume Next
                If Not IsEmpty(CoreXrefSvNm) Then FileCopy ZeusPATH & CoreXrefSvNm, NetworkPath & FileName_PSC & "\1 DIR Files\3 RFE & X-ref Files\" & CoreXrefSvNm
                On Error GoTo 0
            Else
                Set objFolder = objFSO.GetFolder(NetworkPath & FileName_PSC & "\1 DIR Files\3 RFE & X-ref Files\")
                If Not objFolder.Files.Count = 0 Then xrefFlg = 1
                For Each sht In tmWB.Sheets
                    If InStr(UCase(sht.Name), "QC") Then
                        sht.Visible = True
                        sht.Select
                        sht.Unprotect Password = "existentialism"
                        Range("C:C").Find(what:="files posted", lookat:=xlPart).Offset(0, 1).Select
                        If InStr(sht.Name, "MS") Or xrefFlg = 1 Then
                            ActiveCell.Interior.Color = 65280
                            ActiveCell.Offset(0, -1).Interior.ColorIndex = 0
                            sht.Protect Password = "existentialism"
                        Else
                            ActiveCell.Interior.Color = 65535
                            retval = Shell("explorer.exe " & NetworkPath & FileName_PSC & "\1 DIR Files", vbNormalFocus)
                        End If
                        Exit For
                    End If
                Next
            End If
        End If
        
    End If

Application.DisplayAlerts = True
If Not WinOpenFLG = 1 And Not ReviewFlg = 1 Then retval = Shell("explorer.exe " & NetworkPath & FileName_PSC, vbNormalFocus)

QCend:
FormNM.PostToI.ForeColor = &HC0C0C0
FormNM.PostToI.Font.Bold = True

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNetNotSet:
Call SetNetPATH(NetNm)
Resume

errhndlFndExWB:
MsgBox "Please select Initial Extract."
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intchoice = Application.FileDialog(msoFileDialogOpen).Show
 If intchoice <> 0 Then
    wbstr = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        Set exwb = Workbooks.Open(wbstr)
    Else
        wbfoundFLG = 0
    End If
Resume FndExWB

errhndlManualSpot:
Debug.Print c.Row
MsgBox "Cannot evaluate this extract, must spot check manually."
Resume RFsetup

errhndlTRYNAME:
Resume TryName
TryName:
On Error GoTo errhndlNOstd
StdName = Trim(LCase(wfnamerng.Find(what:=orgnlID.Offset(0, 5).Value, lookat:=xlWhole).Offset(0, 1).Value))
GoTo stdMbr

errhndlNOstd:
StdName = LCase(Trim(orgnlID.Offset(0, 5).Value))
Resume stdMbr

End Sub



