Attribute VB_Name = "A__CreateReport"
Sub Create_Report(Optional BulkRun As Boolean)

Dim ReportTotalTime As Long

'set variables
'======================================================================================================================================================
    If Not BulkRun = True Then
        If Not FUN_Save = vbYes Then Exit Sub
    End If
    If ZeusForm.asscSystems.ListCount = 0 And ZeusForm.asscMembers.ListCount = 0 Then
        MsgBox "Please input in the setup tab the systems and/or members you would like to include in your report."
        Exit Sub
    ElseIf ZeusForm.asscContracts.ListCount = 0 Then
        MsgBox "Please input in the setup tab the contracts you would like to include in your report."
        Exit Sub
    End If
    ReportStartTime = Time
    SetupSwitch = FUN_SetupSwitch
    SetupSwitch = 2
    MainCall = 1
    CreateReport = True
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.CalculateBeforeSave = False
    Application.ScreenUpdating = False
    Application.AutoRecover.Enabled = False
    
'Initialize Zeus template
'======================================================================================================================================================
    TempName = "Zeus_Temp" & " (" & NetNm & " - " & FileName_PSC & ").xlsx"
    For Each Wb In Workbooks
        If Wb.Name = TempName Then
            Wb.Close (False)
        End If
    Next
    On Error GoTo ERR_MstrTmplt
    If Dir(TemplatePATH & "\" & MasterTemplate, vbDirectory) = vbNullString Then FileCopy SupplyNetPATH & "\Zeus\Prod\Prod1\Local Components\Templates\" & MasterTemplate, TemplatePATH & "\" & MasterTemplate
    On Error GoTo 0
    FileCopy TemplatePATH & "\" & MasterTemplate, ZeusPATH & "\" & TempName
    Set reportWB = Workbooks.Open(ZeusPATH & "\" & TempName)
    'reportWB.SaveAs (ZeusPATH & "\Zeus_Temp.xlsx")
    'Set tmWB = ReportWB
    
    reportWB.Sheets("Index").Range("C8").Value = Date
    reportWB.Sheets("Index").Range("C7").Value = NetNm & " - " & PSCVar
    
    Call SetBKMRKs  '>>>>>>>>>>
    
    'Hardcode sumproduct formulas
    '---------------------------------
'    Sheets("Vizient Contracts - Conv").Range("J:J").Value = Sheets("Vizient Contracts - Conv").Range("J:J").Value
'    Sheets("line item data").Range("Z:Z").Value = Sheets("line item data").Range("Z:Z").Value

'Search for Spend
'======================================================================================================================================================
    Application.StatusBar = "Searching for Spend Data...Please Wait"
    If Trim(ZeusForm.AsscExtract.Caption) = "" Then
    
        'enter search crit from setup tab
        '--------------------------
        If Not SpendPSCInit = 1 Then
            If Trim(ZeusForm.spendPSC.Value) = "" And Not Trim(ZeusForm.asscPSC.Value) = "" Then
                ZeusForm.spendPSC.Value = Trim(ZeusForm.asscPSC.Value)
                ZeusForm.sPscOR.Value = True
                SpendPSCInit = 1
            End If
        End If
        
        If Not SpendConInit = 1 Then
            If Trim(ZeusForm.spendContract.Text) = "" And Not ZeusForm.asscContracts.ListCount = 0 Then
                For i = 0 To ZeusForm.asscContracts.ListCount - 1
                    ZeusForm.spendContract.Text = ZeusForm.spendContract.Text & ZeusForm.asscContracts.List(i) & "; "
                Next
                ZeusForm.spendContract.Text = Left(ZeusForm.spendContract.Text, Len(ZeusForm.spendContract.Text) - 2)
                ZeusForm.sContractOR.Value = True
                SpendConInit = 1
            End If
        End If
        
        'Search
        '--------------------------
        Call Spend_Search  '>>>>>>>>>>
        If endFLG = 1 Then GoTo EndClean
        DoEvents
        Application.StatusBar = "Formatting Spend Extract..."
        Call StdzExtract  '>>>>>>>>>>
        
    Else
        Call FUN_TestForSheet("Spend Search")
        Cells.Clear
        Workbooks.Open (ZeusPATH & Trim(ZeusForm.AsscExtract.Caption))
        Range("A:AQ, BH:BH").Copy
        tmWB.Sheets("Spend Search").Range("A1").PasteSpecial xlPasteAll
        ActiveWorkbook.Close (False)
        Call StdzExtract(1)  '>>>>>>>>>>
    End If
    
    
'Populate members
'======================================================================================================================================================
    Call RefreshMembers(3)
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    '[To add checkboxes]
    'For i = 60 + suppNMBR To 208
    '    ActiveSheet.CheckBoxes.Add(Range("G10").Offset(i, 0).Left + Range("G10").Offset(i, 0).Width / 3, Range("G10").Offset(i, 0).Top - Range("G10").Offset(i, 0).Height / 32, 10, 10).Select
    '    Selection.Caption = ""
    '    Selection.LinkedCell = Range("A10").Offset(i, 0).Address
    'Next
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'Delete/Hardcode & hide unused supplier sections
'======================================================================================================================================================
    If Not suppNMBR = 10 Then
        Range(prsBKMRK.Offset(0, suppNMBR * 2 + 1), prsBKMRK.Offset(MbrNMBR + 2, 20)).Delete Shift:=xlToLeft
        Range(MSGraphBKMRK.Offset(0, suppNMBR * 2 + 1), MSGraphBKMRK.Offset(MbrNMBR + 2, 20)).Delete Shift:=xlToLeft
        Set ConTblBKMRK = Sheets("Index").Range("C:C").Find(what:="Supplier Name", lookat:=xlWhole)
        Range(ConTblBKMRK.Offset(0, suppNMBR + 1), ConTblBKMRK.Offset(60, 10)).Delete Shift:=xlToLeft
        For i = 1 To 10 - suppNMBR
            Set tblBkmk = Sheets("Vizient Contracts - NC").Range("A1590").End(xlUp)
            Range(tblBkmk.Offset(0, 2), tblBkmk.Offset(-MbrNMBR + 2, 18)).ClearContents
            Range(tblBkmk.Offset(4, 0), tblBkmk.Offset(-MbrNMBR - 3, 18)).EntireRow.Hidden = True
            Set tblBkmk = Sheets("Vizient Contracts - Conv").Range("A1590").End(xlUp)
            Range(tblBkmk.Offset(0, 2), tblBkmk.Offset(-MbrNMBR + 2, 20)).ClearContents
            Range(tblBkmk.Offset(4, 0), tblBkmk.Offset(-MbrNMBR - 3, 20)).EntireRow.Hidden = True
        Next
        
'        Sheets("Vizient Contracts - NC").Range("B" & 8 + suppNMBR * (MbrNMBR + 8) & ":B1587").EntireRow.Value = Sheets("Vizient Contracts - NC").Range("B" & 8 + suppNMBR * (MbrNMBR + 8) & ":B1587").EntireRow.Value
'        Sheets("Vizient Contracts - NC").Range("B" & 8 + suppNMBR * (MbrNMBR + 8) & ":B1587").EntireRow.Hidden = True
'        Sheets("Vizient Contracts - Conv").Range("B" & 8 + suppNMBR * (MbrNMBR + 8) & ":B1587").EntireRow.Value = Sheets("Vizient Contracts - Conv").Range("B" & 8 + suppNMBR * (MbrNMBR + 8) & ":B1587").EntireRow.Value
'        Sheets("Vizient Contracts - Conv").Range("B" & 8 + suppNMBR * (MbrNMBR + 8) & ":B1587").EntireRow.Hidden = True
    End If
    
    
'Import tier info
'=================================================================================================================
    DoEvents
    Application.StatusBar = "Importing Tier Info..."
    Call Import_TierInfo  '>>>>>>>>>>

'Import pricefiles
'=================================================================================================================
    DoEvents
    Application.StatusBar = "Importing Pricefiles..."
    Call Import_Pricefile  '>>>>>>>>>>

'import UNSPSC
'====================================================================================================
    DoEvents
    Application.StatusBar = "Importing UNSPSC..."
    Call Import_UNSPSC  '>>>>>>>>>>
    
'populate admin fees
'====================================================================================================
    DoEvents
    Application.StatusBar = "Importing Admin fees..."
    Call Import_AdminFees    '>>>>>>>>>>

'Populate Notes tab
'====================================================================================================
    Sheets("Notes").Range("D1").Value = Usr
    Sheets("Notes").Range("D2").Value = ZeusForm.plvlRngSet.Value
    Sheets("Notes").Range("D3").Value = ZeusForm.SuppRngSet.Value
    Sheets("Notes").Range("D4").Value = ZeusForm.bnchRngSet.Value
    Range("D:D").EntireColumn.AutoFit
    Range("D:D").HorizontalAlignment = xlLeft

'import xref
'====================================================================================================
    If Not ZeusForm.PartialChk = True Then
        If Not Trim(ZeusForm.AsscXref.Caption) = "" Then
            If Not xrefopt = vbNo Then    'if user chose not to use an xref then skip tiermax section
                Call ExtractCore          '>>>>>>>>>>
                'Call METH_ExtractIntelli '>>>>>>>>>>
                Call FormatCrossRefTabs   '>>>>>>>>>>
            End If
        End If
    End If
    
'Import spend
'====================================================================================================
If Trim(ZeusForm.AsscExtract.Caption) = "" Then
    Sheets("Spend Search").Select
    
    Call Import_Spend
    
    Sheets("Spend search").Cells.Clear
    Application.StatusBar = False
Else
    Call Import_Spend  '>>>>>>>>>>
End If

'Import Benchmarking
'====================================================================================================
    Call Import_Benchmarking   '>>>>>>>>>>
    
'[VARIANCES]
'====================================================================================================================================
    '[TBD]When sorting, the top row of formulas needs to stay at the top or be entered at the end of the report being created

    If ZeusForm.ZeusResolveVar = True And Not CalcChkFLG = vbNo Then
        Sheets("line item data").Select
    
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'Resolve price leveling variances
    '================================================================================================================================
        'Calculating.Show (False)

'        plvlVarBkmrk = Rows("2:2").Find(what:="Price Level Each Cost", lookat:=xlPart).Address
'        Set benchbkmrk = Sheets("Line Item Data").Rows("2:2").Find(what:="10th Variance (%)", lookat:=xlWhole)
'        calcstr = "U:U, V:V, X:X, " & Range(plvlVarBkmrk).EntireColumn.Address & ", " & Range(plvlVarBkmrk).Offset(0, 1).EntireColumn.Address & ", " & Range(plvlVarBkmrk).Offset(0, 3).EntireColumn.Address & ", " & Range(plvlVarBkmrk).Offset(0, 4).EntireColumn.Address & ", " & Range(plvlVarBkmrk).Offset(0, 5).EntireColumn.Address & ", " & benchbkmrk.EntireColumn.Address & ", " & benchbkmrk.Offset(0, -2).EntireColumn.Address & ", " & benchbkmrk.Offset(0, -3).EntireColumn.Address
'        For I = 0 To suppNMBR - 1
'            calcstr = calcstr & ", " & Range(Range("AM1").Offset(0, I * 18), Range("AM1").Offset(0, I * 18 + 11)).EntireColumn.Address
'        Next
'        Sheets("Line Item Data").Range(calcstr).Calculate
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Application.StatusBar = "Calculating Line Item Data...Please Wait"
        Sheets("Line Item Data").Calculate
        Application.StatusBar = False
        'Unload Calculating
        Call StrtEinstein '>>>>>>>>>>
    
    End If

'Add suppliers
'=======================================================================================
    If Not LargeReport = True Then Call AddSuppliers  '>>>>>>>>>>

'Import PRS
'=======================================================================================
    Call Import_PRS   '>>>>>>>>>>>

'Insert QC sheet
'=======================================================================================
    Call QC_Template  '>>>>>>>>>>


'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'[CLEAN AND END]//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 
'Finalize & format
'===================================================================================================
     
    'delete added tabs
    '------------------------
    On Error Resume Next
    Sheets("Supplier Pivot").Visible = False
    Sheets("ClinicalQC").Delete
    Sheets("xxStdNames").Visible = False
    Sheets("xxCalculations").Delete
    Sheets("Scopeguide").Visible = False
    Sheets("Spend Search").Delete
    On Error GoTo 0
    
    'Freeze top row on line item data
    '------------------------
    Sheets("Line item data").Select
    Rows("5:5").Select
    ActiveWindow.FreezePanes = True

'Calculate
'=======================================================================================
    If Not CalcChkFLG = vbNo Then
        'Calculating.Show (False)
        DoEvents
        Application.StatusBar = "Please wait while your report is calculating and finalizing..."
        
        'calculate
        '---------------------------------
        reportWB.SaveAs (ZeusPATH & "\" & NetNm & "_" & FileName_PSC & "_S2TM_" & Replace(Sheets("index").Range("C8").Value, "/", "-") & ".xlsx")
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
        Application.StatusBar = False
                
        'Add sumproduct formulas back and calculate
        '---------------------------------
'        Application.StatusBar = "Caluclating Unique Product formula..."
'        For i = 1 To suppNMBR
'            SuppCol = Sheets("line item data").Range("BG5:BG99999").Offset(0, (i - 1) * 30).Address
'            suppoffset = (MbrNMBR + 8) * (i - 1)
'            ConvBKMRK.Offset(suppoffset + 1, 8).Formula = "=CONCATENATE(IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$P$5:$P$99999=$B12)*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!Z$5:Z$99999)),DOLLAR(SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$P$5:$P$99999=$B12)*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!AJ$5:AJ$99999)),0)),"" of "",IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUMIF('Line Item Data'!$P:$P,$B12,'Line Item Data'!$Z:$Z),DOLLAR(SUMIF('Line Item Data'!$P:$P,$B12,'Line Item Data'!$AJ:$AJ),0)))"
'            ConvBKMRK.Offset(suppoffset + 1, 8).AutoFill Destination:=Range(ConvBKMRK.Offset(suppoffset + 1, 8), ConvBKMRK.Offset(suppoffset + MbrNMBR, 8))
'            ConvBKMRK.Offset(suppoffset + MbrNMBR + 1, 8).Formula = "=CONCATENATE(IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!Z$5:Z$99999)),DOLLAR(SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!AJ$5:AJ$99999)),0)),"" of "",IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUM('Line Item Data'!$Z:$Z),DOLLAR(SUM('Line Item Data'!$AJ:$AJ),0)))"
'        Next
'        Sheets("Line item data").Range("Z5").Formula = "=IF(SUMPRODUCT(($X$5:$X5=$X5)*($P$5:$P5=$P5))>1,0,1)"
'        Sheets("Line item data").Range("Z5").AutoFill Destination:=Range(Sheets("Line item data").Range("Z5"), Sheets("Line item data").Range("A4").End(xlDown).Offset(0, 25))
'        Application.Calculation = xlCalculationAutomatic
'        Application.Calculation = xlCalculationManual
'        Application.StatusBar = False
    End If

'Save new file
'=======================================================================================
    Application.DisplayAlerts = False
    DoEvents
    Application.StatusBar = "Please wait, Saving new Report..."
    'reportWB.Save
    Application.DisplayAlerts = True
   
'import remaining spend items if large report
'=======================================================================================
    If LargeReport = True Then
        
        'Search for remainder
        '--------------------------
        LargeReport = False
        CreateReport = False
        Call Spend_Search  '>>>>>>>>>>
        CreateReport = True
        If endFLG = 1 Then GoTo EndClean
        DoEvents
        Application.StatusBar = "Formatting Spend Extract..."
        Call StdzExtract  '>>>>>>>>>>
        Sheets("Spend Search").Select
    
        Call Import_Spend
        
        Sheets("Spend search").Cells.Clear
        Application.StatusBar = False
        Call AddSuppliers  '>>>>>>>>>>
        
    End If
   
   
'Run QC
'=======================================================================================
    QCform.Show (False)
    QCform.ScrollTop = 0
    Call QC_Main
    
'Delete temp
'=======================================================================================
    On Error Resume Next
    Kill ZeusPATH & "\" & TempName
    On Error GoTo 0
    
'calculate time elapsed
'=======================================================================================
    ReportEndTime = Time
    ReportTotalTime = DateDiff("s", ReportStartTime, ReportEndTime)
    If ReportTotalTime > 60 Then
        mintime = Int(ReportTotalTime / 60)
        secTime = ReportTotalTime Mod 60
        timestr = mintime & "." & secTime & " minutes"
        If mintime > 60 Then
            hrtime = Int(mintime / 60)
            mintime = mintime Mod 60
            timestr = hrtime & "." & mintime & " hours"
        End If
    Else
        timestr = ReportTotalTime & " seconds"
    End If
    
'Run QC
'=======================================================================================
    CreateReport = False
    
'    If Not ZeusForm.PartialChk = True Then
'        If Not CalcChkFLG = vbNo Then
'            DoEvents
'            Application.StatusBar = "Running QC..."
'            QCform_S2TM.Show (False)
'            QCform_S2TM.ScrollTop = 0
'            Call QC_Main
'        End If
'    End If '(END IF: partial report)
    
    'Clean up
    '----------------
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Sheets("Vizient Contracts - NC").Select
    Range("A1").Select
    Sheets("Vizient Contracts - Conv").Select
    Range("A1").Select
    Sheets("Index").Select
    Range("A1").Select
    
    MsgBox ("Total time elapsed: " & timestr)
    
    If Not Trim(ZeusForm.AsscExtract.Caption) = "" Then
        SourceVar = "User Extract"
    ElseIf ZeusForm.NRSrch.Value = True Then
        SourceVar = "Network Run"
    ElseIf ZeusForm.extractSrch.Value = True Then
        SourceVar = "Extract"
    Else
        SourceVar = "RDM"
    End If
    
    ttlHCO = Application.CountA(tmWB.Sheets("Line Item Data").Range("A:A")) - 3
    eBODY = "<p><b><u>Tier Max(" & SourceVar & ")</u></b></p>" & _
            "Total Runtime: " & timestr & _
            "<br>" & "Lines: " & ttlHCO & _
            "<br>" & "Suppliers: " & suppNMBR & _
            "<br>" & "Members: " & MbrNMBR & _
            "<br>" & "Total Spend: " & Format(Sheets("Line Item Data").Range("AJ3").Value, "$#,##0.0") & _
            "<p><b>" & NetNm & "-" & PSCVar & "</b></p>" & _
            Usr & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            "<br>" & _
            ThisWorkbook.Name

    Call MailMetrics("Jason.Solberg@vizientinc.com", "Metrics: Tier Max", eBODY, "bforrest@novationco.com", "", 1)




Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoWB:
Workbooks.Add
Resume

ERR_MstrTmplt:
MsgBox "Unable to download updated master template.  Please copy " & MasterTemplate & " from network folder here: " & vbCrLf & vbCrLf & SupplyNetPATH & "\Zeus\Prod\Prod1\Local Components\Templates " & vbCrLf & vbCrLf & "to local Template folder here: " & vbCrLf & vbCrLf & TemplatePATH
Exit Sub

errhndlefilter:
Rows("2:2").AutoFilter
On Error GoTo 0
Resume

EndClean:
endFLG = 0
On Error Resume Next
reportWB.Close (False)
Kill ZeusPATH & "\" & TempName
On Error GoTo 0
Exit Sub

'errhndlAllNames:
'Resume 88
'88  On Error Resume Next
'    Set AllNames = Application.InputBox(prompt:="Select range of standardized names to be included in analysis, otherwise cancel.", Title:="SPECIFY RANGE", Type:=8)
'    On Error GoTo 0
'    GoTo 8
'
End Sub


