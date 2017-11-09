VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZeusForm 
   Caption         =   "Zeus"
   ClientHeight    =   11280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   24720
   OleObjectBlob   =   "ZeusForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ZeusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IndivTitleArray(1 To 23) As String
Dim IndivDescArray(1 To 23) As String
Dim PSCArray() As String
Dim PSCPath() As String
Dim ToolsArray(1 To 7) As String
Dim ToolsPath(1 To 7) As String
Dim CommonArray(1 To 11) As String
Dim CommonPath(1 To 11) As String
Dim TemplateArray(1 To 6) As String
Dim TemplateMETH(1 To 6) As String
Dim ReportPath() As String
Dim CUnmbr As Integer
Dim CUremoves() As New ZeusEvents
Dim EinsteinInitFLG As Integer
Dim FolderInitFlg As Integer
Dim TMinitFLG As Integer
Dim ReportInitflg As Integer
Dim PrevNetVal As String
Dim AddNoteFormFLG As Integer
Dim SetupTabInit As Integer
Dim GeneralWidthVar As Integer
Dim GeneralHeightVar As Integer
Dim NoRec As Integer

'Public WithEvents XLApp As Excel.Application
'Dim mXLHwnd As Long    'Excel's window handle
'Dim mhwndForm As Long  'The userform's window handle
'Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Const GWL_HWNDPARENT As Long = -8
'
'Private Sub XLApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
'    If Val(Application.Version) >= 15 And mhwndForm <> 0 Then  'Basear o form na janela ativa do Excel.
'        mXLHwnd = Application.hwnd    'Always get because in Excel 15 SDI each wb has its window with different handle.
'        SetWindowLongA mhwndForm, GWL_HWNDPARENT, mXLHwnd
'        SetForegroundWindow mhwndForm
'    End If
'End Sub
'
'Private Sub XLApp_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
'    If Not Me.Visible Then Me.Show vbModeless
'End Sub


Private Sub MemberExport_Click()

Sheets.Add
For i = 0 To MembersReturned.ListCount - 1
    Range("A1").Offset(i, 0).Value = MembersReturned.List(i)
Next

End Sub
Private Sub SystemExport_Click()

Sheets.Add
For i = 0 To SystemsReturned.ListCount - 1
    Range("A1").Offset(i, 0).Value = SystemsReturned.List(i)
Next

End Sub

'name each doc or maybe folder with date at end that way can compare easily and only import if dif dates. or maybe compare modified dates to see if file has been modified since last cache date.  maybe have a log file listing each of the import dates for each file and check it to see which ones need to be imported
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'[Top form]////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub UserForm_Initialize()

Dim conn As New ADODB.Connection
Dim recset As New ADODB.Recordset

'set variables if not already set
'===============================================================================================
On Error Resume Next
ZeusTitle = Replace(ThisWorkbook.Name, "Zeus_App(", "")
ZeusForm.Caption = "Zeus v" & Replace(ZeusTitle, ").xlam", "")

Application.DisplayAlerts = False
Application.StatusBar = "Starting Zeus...Please wait"

'If Val(Application.Version) >= 15 Then        'Only makes sense on Excel 2013 and up
'    Set XLApp = Application
'    mhwndForm = FindWindowA(ZeusForm.Caption, Caption)
'End If

'Load Keyboard Shortcuts
'===============================================================================================

'Workbook shortcuts
'---------------------------------
'Application.OnKey "^+W", "goWFwb"
Application.OnKey "^+W", "goTMwb"

'tab shortcuts
'---------------------------------
'Application.OnKey "%+!", "GoNotes"
'Application.OnKey "%+1", "GoQC" '(if exists)
Application.OnKey "^+Q", "GoQC"
Application.OnKey "^+>", "GoLID"
Application.OnKey "^+<", "GoIndex"
'Application.OnKey "%+C", "GoCMSg"
'Application.OnKey "^H", "GoHCO"
'Application.OnKey "%+B", "GoBMP"
'Application.OnKey "%+R", "GoItemsRemoved"

'tool Shortcuts
'---------------------------------
Application.OnKey "^+P", "PricefileAdjustments_Indiv"
Application.OnKey "^+D", "HighlightDups"
Application.OnKey "^+U", "Uppercase"
Application.OnKey "^+N", "TxtToCol"
Application.OnKey "^+B", "AddBorders"
Application.OnKey "^+C", "calcSelection"
Application.OnKey "^+R", "RemoveIOC"
Application.OnKey "^+A", "addIOC"
Application.OnKey "%+{left}", "DeleteLeft"
Application.OnKey "%+{Up}", "DeleteUp"
Application.OnKey "^+S", "NormalizeData"
Application.OnKey "^+V", "CPvalues"
Application.OnKey "^+F", "AddHighlight"

'Sherlock shortcuts
'---------------------------------
Application.OnKey "^+I", "ClinicalIS"
Application.OnKey "^+O", "ClinicalOOS"
Application.OnKey "^+T", "ClinicalTBD"

'Setup variables
'===============================================================================================
Call SetPathVariables
Call SetNetsArray
'Call SetBKMRKs
'SetupSwitch = FUN_SetupSwitch(1)

'Pull psc list from edb
'==========================================================================================================================================================
Application.ScreenUpdating = False
On Error GoTo nowb
Sheets.Add

On Error GoTo errhndlNORECSET
Application.StatusBar = "Starting Zeus...Importing PSCs"
conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
sqlstr = "SELECT DISTINCT ATTRIBUTE_VALUE_NAME FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL WHERE ATTRIBUTE_NAME = 'PRODUCT SUB-CATEGORY' AND STATUS_KEY = 'ACTIVE' AND ATTRIBUTE_VALUE_STATUS = 'A' AND not ATTRIBUTE_VALUE_NAME = '' ORDER BY ATTRIBUTE_VALUE_NAME"
recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
Range("A1").CopyFromRecordset recset
recset.Close

pscNmbr = Application.CountA(Range(Range("A1"), Range("A1").End(xlDown)))
For i = 1 To pscNmbr
    ZeusForm.asscPSC.AddItem Range("A1").Offset(i - 1, 0).Value
    Application.StatusBar = "Starting Zeus...Please Wait (PSCs: " & pscNmbr - i & ")"
Next
Range(Range("A1"), Range("A1").End(xlDown)).ClearContents

'pull network list from EDB
'=========================================================================================================================================================
For itm = 1 To UBound(NtwkNmArray)
    asscNetwork.AddItem NtwkNmArray(itm)
Next

sqlstr = "SELECT DISTINCT networkname FROM MEMT1MEINQ WHERE not networkname = '' ORDER BY networkname" 'WHERE not networkname = ''"
recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

Range("A1").CopyFromRecordset recset
recset.Close

NTWKnmbr = Application.CountA(Range(Range("A1"), Range("A1").End(xlDown)))
For i = 1 To NTWKnmbr
    Aitm = Trim(Range("A1").Offset(i - 1, 0).Value)
    For j = 1 To NetNmbr
        If Aitm = NtwkEDBArray(j) Then GoTo nxtNtwk
    Next
    ZeusForm.asscNetwork.AddItem Aitm
'    netcnt = netcnt + 1
'    ReDim Preserve NtwkIDArray(1 To NetNmbr + netcnt)
'    NtwkIDArray(NetNmbr + netcnt) = Trim(Recset.Fields(1))
nxtNtwk:
    Application.StatusBar = "Starting Zeus...Please Wait (Networks: " & NTWKnmbr - i & ")"
Next
Application.StatusBar = "Starting Zeus...Please Wait (Importing Company Codes)"

sqlstr = "SELECT DISTINCT companyid FROM MEMT1MEINQ"
recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

For i = 1 To recset.RecordCount
    If Not Trim(recset.Fields(0)) = "" Then ZeusForm.AsscCompany.AddItem Trim(recset.Fields(0))
    recset.MoveNext
Next
ZeusForm.AsscCompany.Value = "001"

Set recset = Nothing
Set conn = Nothing

If tempwb = 1 Then
    ActiveWorkbook.Close (False)
Else
    ActiveSheet.Delete
End If

Application.StatusBar = False

'establish initial connection to RDM
'===============================================================================================
If RDMconn = "" Then
    On Error Resume Next
    connstr = "Driver={Microsoft ODBC for Oracle};CONNECTSTRING=" & RDMConnStr
    RDMconn.Open connstr
    On Error GoTo 0
End If

'Set HxW
'===============================================================================================
NRSrch.Value = False
extractSrch.Value = False
GeneralWidthVar = 65
GeneralHeightVar = 185

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height / 2 - Me.Height / 4
Me.Left = Application.Left + Application.Width / 2 - Me.Width / 4

'ZeusPages.Height = 315
'ZeusPages.Width = 550
'ZeusForm.Height = ZeusPages.Height
'ZeusForm.Width = ZeusPages.Width

ZeusPages.Value = 7
ZeusPages.Value = 0 '(show setup page on startup)

'AddToForm MIN_BOX
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNORECSET:
Resume CloseConn
CloseConn:
On Error Resume Next
Set recset = Nothing
Set conn = Nothing
Application.StatusBar = False
ConnNotFnd = 1
Exit Sub

nowb:
Workbooks.Add
tempwb = 1
Resume Next



End Sub
Private Sub UserForm_Activate()

'Call AddMinBttn(Me.Caption)
AddToForm MIN_BOX

End Sub
Private Sub userform_terminate()

''if ZeusForm closes close all other forms
''--------------------------
''(maybe just use "End" function? No cause i still want StatStracker to run)
'On Error Resume Next
'If CurrentProject.AllForms("addnoteform").IsLoaded Then Unload AddNoteForm
'If RubiksForm.Visible = True Then Unload RubiksForm
'Unload HermesInitialPrompt
'Unload QCform_DP2
'Unload QCform_MS
'Unload QCform_S2TM
'Unload ReportCheck
'Unload SherlockForm
SysCNT = 0


End Sub
Sub ZeusPages_Change()

If ZeusNotes = 1 Then
    If AddNoteForm.Left = ZeusForm.Left + 80 And AddNoteForm.Top = ZeusForm.Top + 35 Then Unload AddNoteForm
    ZeusNotes = 0
End If

'Setup
'=================================================================================================
'=================================================================================================
If ZeusPages.SelectedItem.Index = 0 Then
    
    If Not SetupTabInit = 1 Then
        
        'populate lookup lists
        '----------------------------
        For i = 0 To asscPSC.ListCount - 1
            ContractPSCcrit.AddItem asscPSC.List(i)
        Next
        For i = 0 To asscNetwork.ListCount - 1
            SystemNetworkCrit.AddItem asscNetwork.List(i)
        Next
        'If Not Trim(asscNetwork.Value) = "" Then Call populateResults("MemberSystemCrit", asscNetwork.Value)
        
        'AsscEndDate.Value = Date
        'asscStartDate.Value = Left(Date, Len(Date) - 1) & Right(Date, 1) - 1
        
        SetupTabInit = 1
    
    End If
    
    'find network, PSC, contracts and Date Range
    '----------------------------
    If Trim(asscNetwork.Text) = "" Or Trim(asscPSC.Text) = "" Or asscContracts.ListCount = 0 Then
        SetupSwitch = FUN_SetupSwitch(1)
    Else
        SetupSwitch = FUN_SetupSwitch
    End If
    
    'find assc files
    '----------------------------
    If Not PSCVar = "" Then Call Find_Xref_Extract
    
    'Adjust form HxW
    '----------------------------
    If MemberLookupFrame.Visible = True Then
        ZeusPages.Width = MemberLookupFrame.Left + MemberLookupFrame.Width + GeneralWidthVar
    ElseIf SystemLookupFrame.Visible = True Then
        ZeusPages.Width = SystemLookupFrame.Left + SystemLookupFrame.Width + GeneralWidthVar
    ElseIf ContractLookupFrame.Visible = True Then
        ZeusPages.Width = ContractLookupFrame.Left + ContractLookupFrame.Width + GeneralWidthVar
    Else
        ZeusPages.Width = asscContractFrame.Left + asscContractFrame.Width + GeneralWidthVar + 5
    End If
    
    ZeusPages.Height = AsscMemberFrame.Top + AsscMemberFrame.Height + 10 '+ GeneralHeightVar
    ZeusForm.Height = ZeusPages.Height + 20
    ZeusForm.Width = ZeusPages.Width

'Import Spend
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 1 Then
    
    SetupSwitch = FUN_SetupSwitch
    
    'populate PSC drop down on import spend tab
    '--------------------------
    If spendPSC.ListCount = 0 Then
        For i = 0 To asscPSC.ListCount - 1
            spendPSC.AddItem asscPSC.List(i)
        Next
    End If
    
    'enter values from setup into import spend tab if blank
    '--------------------------
    If Not SpendPSCInit = 1 Then
        If Trim(spendPSC.Value) = "" And Not Trim(asscPSC.Value) = "" Then
            spendPSC.Value = Trim(asscPSC.Value)
            sPscOR.Value = True
            SpendPSCInit = 1
        End If
    End If
    
    If Not SpendConInit = 1 Then
        If Trim(spendContract.Text) = "" And Not asscContracts.ListCount = 0 Then
            For i = 0 To asscContracts.ListCount - 1
                spendContract.Text = spendContract.Text & asscContracts.List(i) & "; "
            Next
            spendContract.Text = Left(spendContract.Text, Len(spendContract.Text) - 2)
            sContractOR.Value = True
            SpendConInit = 1
        End If
    End If
    
    On Error Resume Next
    ZeusForm.PotSpend.Caption = "Potential Spend:  " & Format(Sheets("Line Item Data").Range("X1").Value + Application.sum(Range(Range("W2"), Range("A1").End(xlDown).Offset(0, 22))), "$#,##0.00")
    ZeusForm.PotRows.Caption = "Potential Rows:    " & Range(Sheets("Line Item Data").Range("A3"), Sheets("Line Item Data").Range("A3").End(xlDown)).Count + Range(Range("A2"), Range("A1").End(xlDown)).Count - 1
    On Error GoTo 0
    
    ZeusPages.Height = spendSearch.Top + spendSearch.Height + 45 'GeneralHeightVar
    ZeusPages.Width = spendDiscard.Left + spendDiscard.Width + GeneralWidthVar + 10
    ZeusForm.Height = ZeusPages.Height
    ZeusForm.Width = ZeusPages.Width

'TierMax
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 2 Then
    
    If Not TMinitFLG = 1 Then
    
        'Set Feature names
        '--------------------------
        IndivTitleArray(1) = "Import Tier Information"
        IndivTitleArray(2) = "Import Pricefile"
        IndivTitleArray(3) = "Import Multiple"
        IndivTitleArray(4) = "Refresh Suppliers"
        IndivTitleArray(5) = "Import CheatSheet"
        IndivTitleArray(6) = "Import SNA Sheet"
        IndivTitleArray(7) = "Import Scopeguide"
        IndivTitleArray(8) = "Import UNSPSC"
        IndivTitleArray(9) = "Import PRS"
        IndivTitleArray(10) = "Import Alt PRS"
        IndivTitleArray(11) = "Import Admin Fees"
        IndivTitleArray(12) = "Import Benchmarking"
        IndivTitleArray(13) = "Import CoreXref"
        IndivTitleArray(14) = "Import Intellisource"
        IndivTitleArray(15) = "Format Xrefs"
        IndivTitleArray(16) = "Refresh Members"
        IndivTitleArray(17) = "Standardize Manufacturers"
        IndivTitleArray(18) = "Novaplus"
        IndivTitleArray(19) = "Add Suppliers"
        IndivTitleArray(20) = "Finalize"
        IndivTitleArray(21) = "Create DATxref"
        IndivTitleArray(22) = "Keyword Generator"
        IndivTitleArray(23) = "Import Dates"
        
        'Set Feature Descriptions
        '--------------------------
        IndivDescArray(1) = "Import MPP manufacturer names, contract start/end date, novaplus/standardization program, portfolio executive, and tier information."
        IndivDescArray(2) = "Import individually specified pricefile."
        IndivDescArray(3) = "Import Pricefiles for contracts in contract list."
        IndivDescArray(4) = "Refresh Supplier data with selected contracts."
        IndivDescArray(5) = "Pull cheatsheet for associated PSC from I: drive if exists."
        IndivDescArray(6) = "Import network SNA standardization tab from standardization file. "
        IndivDescArray(7) = "Import data from scopeguide database for PSC if exists."
        IndivDescArray(8) = "Import UNSPSC data."
        IndivDescArray(9) = "Import PRS data for contracts in contract list."
        IndivDescArray(10) = "Import PRS data for alternate contracts."
        IndivDescArray(11) = "Import Admin Fee data directly from the EDB database, pull over items from pricefile tabs, and populate net sales and novaplus fees for each contract."
        IndivDescArray(12) = "Import benchmarking data associated with the PSC, contract numbers, and look up each catalog number on the Line Item Data tab if not already found."
        IndivDescArray(13) = "Import xref data from the DATxref file specified on the Zeus setup tab.  Xref data must be on a tab labeled ""Core"" and in Catalog number/Description format."
        IndivDescArray(14) = "Import intellisource data from downloaded intellisource files in the Zeus folder.  Each contract must be in it's own separate file."
        IndivDescArray(15) = "Format xref tabs by populating vlookups, formatting headers, standardizing cell formatting, removing items that don't cross, and sorting by catalog number > source > price."
        IndivDescArray(16) = "Get current member standardization and adjust tables to fit current members."
        IndivDescArray(17) = "Create a pivot table comparing manufacturer names to catalog numbers and standardize manufacturers if more than one per given catalog number."
        IndivDescArray(18) = "Based on tier info table, check catalog numbers in corresponding Novaplus supplier sections on the Line Item Data tab and convert to Novaplus catalog numbers from Novaplus file in Zeus folder."
        IndivDescArray(19) = "Add non contracted suppliers to the MarketShare table until the total spend for all others is <5% or until total suppliers is 10."
        IndivDescArray(20) = "Save report to its corresponding initiative folder on the I: drive, upload original extract, ASF extract, pricefiles, DATxref, and CoreXref to their respective folders, and pull out PIM data marked for validation."
        IndivDescArray(21) = "Look up Xref data in the Xref database and format Core xref files for import to TierMax Report."
        IndivDescArray(22) = "Get frequency for single word, 2 word, and 3 word phrases of item descriptions in pricefies."
        IndivDescArray(23) = "Import date ranges for member spend in dataset."

        For Each METH In IndivTitleArray
            IndivSelect.AddItem METH
        Next
            
        'populate QC
        '----------------------
        With QCReviewSelect
            .AddItem "QC"
            .AddItem "Review"
        End With
        With QCtypeSelect
            .AddItem "S2TM"
            '.AddItem "MS"
            '.AddItem "DP2"
        End With
        QCReviewSelect.Value = "QC"

        TMinitFLG = 1
        
    End If
    
    ZeusForm.ZeusPages.Height = DescWindow.Top + DescWindow.Height + 5 'GeneralHeightVar
    ZeusForm.ZeusPages.Width = IndivGo.Left + IndivGo.Width + GeneralWidthVar + 3
    ZeusForm.Height = ZeusPages.Height + 20
    ZeusForm.Width = ZeusPages.Width '+ 3
    
'Sherlock
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 3 Then
    
    ZeusPages.Height = GeneralHeightVar
    ZeusPages.Width = Sherlockwidthbkmrk.Left + Sherlockwidthbkmrk.Width + GeneralWidthVar
    ZeusForm.Height = ZeusPages.Height
    ZeusForm.Width = ZeusPages.Width

'Rubiks
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 4 Then
    
    ZeusPages.Height = PFchkBox.Top + PFchkBox.Height + 5
    'ZeusPages.Width = SandSFrame.Left + SandSFrame.Width + GeneralWidthVar
    ZeusPages.Width = ZeusForm.PkgMismtchLabel.Left + ZeusForm.PkgMismtchLabel.Width + GeneralWidthVar
    ZeusForm.Height = ZeusPages.Height + 20
    ZeusForm.Width = ZeusPages.Width
    If Not EinsteinInitFLG = 1 Then
        With plvlRngSet
            For i = 1 To 20
                .AddItem Format(i * 0.05, "0%")
            Next
            .Value = Format(0.5, "0%")
        End With
        With SuppRngSet
            For i = 1 To 20
                .AddItem Format(i * 0.05, "0%")
            Next
            .Value = Format(0.5, "0%")
        End With
        With bnchRngSet
            For i = 1 To 20
                .AddItem Format(i * 0.05, "0%")
            Next
            .Value = Format(0.9, "0%")
        End With
        With msmtchSelect
            For i = 1 To 20
                .AddItem Format(i * 0.05, "0%")
            Next
            .Value = Format(0.65, "0%")
        End With
        CrawlerChkBox.Value = True
        
        '[TBD]populate initial Einstein stats
        EinsteinInitFLG = 1
    End If

'Notes
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 5 Then
    On Error GoTo errhndlePreReport
        
    'Make sure TMwb is active and variables are set
    '---------------------------------
    SetupSwitch = FUN_SetupSwitch
    If FUN_SetupSwitch = 2 Then
        For Each sht In ActiveWorkbook.Sheets
            If sht.Name = "Notes" Then
                ZeusNotes = 1
                AddNoteForm.Show (False)
                Exit For
            End If
        Next
    Else
        ZeusNotes = 0
        GoTo PreReport
    End If
    
PreReport:
    If ZeusNotes = 0 Then
        On Error Resume Next
        Unload AddNoteForm
        NotesOpen = 0
        On Error GoTo 0
        NotesUnavailable.Visible = True
        ZeusPages.Width = 200
    Else
        NotesUnavailable.Visible = False
        AddNoteForm.Left = ZeusForm.Left + 80
        AddNoteForm.Top = ZeusForm.Top + 35
        ZeusPages.Width = AddNoteForm.Width + 100
    End If

    ZeusPages.Height = GeneralHeightVar
    ZeusForm.Height = ZeusPages.Height
    ZeusForm.Width = ZeusPages.Width
    
'Folders
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 6 Then
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not FolderInitFlg = 1 Then
        
        'populate networks
        '---------------------
        For i = 1 To UBound(NtwkNmArray)
            FoldersNtwkList.AddItem NtwkNmArray(i)
        Next
       
        'populate Common files
        '---------------------
        ToolsArray(1) = "SNAP"
        ToolsArray(2) = "SNA Standardization"
        ToolsArray(3) = "PRS"
        ToolsArray(4) = "UNSPSC"
        ToolsArray(5) = "Scopeguide"
        ToolsArray(6) = "Novaplus"
        'ToolsArray(7) = "Benchmarking"
        ToolsArray(7) = "Xref DB"
        
        ToolsPath(1) = SupplyNetPATH & "\Analytics\SNAP\SNAP 2.4\SNAP_2.4.accdb"
        ToolsPath(2) = Stdzn_Index_PATH & "\" & NetNm & ".xlsx"
        ToolsPath(3) = SupplyNetPATH & "\Analytics\Analytical Tools\Pulling PRS Spend\PRS Generator.accdb"
        ToolsPath(4) = SupplyNetPATH & "\Analytics\Analytical Tools\Product Segmentation.accdb"
        ToolsPath(5) = SupplyNetPATH & "\Analytics\DAT Resources\Scope Guide.accdb"
        ToolsPath(6) = ZeusPATH & "\1-DB shortcuts\NOVAPLUS_NonPharmacy_ProductList.xlsx"
'        Set objFolder = objFSO.GetFolder(BenchPATH)
'        For Each ofile In objFolder.Files
'            If InStr(LCase(ofile.Name), "benchmark") And Not InStr(LCase(ofile.Name), "~$") Then ToolsPath(7) = ofile.Path
'        Next

        ToolsPath(7) = XrefdbPATH
        
        For itm = 1 To UBound(ToolsPath)
            FoldersToolsList.AddItem ToolsArray(itm)
        Next
        
        'populate Common folders
        '---------------------
        CommonArray(1) = "3 Day Temp"
        CommonArray(2) = "Supply Networks"
        CommonArray(3) = "DAT resources"
        CommonArray(4) = "Extracts"
        CommonArray(5) = "Cheatsheets"
        CommonArray(6) = "Zeus"
        CommonArray(7) = "Benchmark"
        CommonArray(8) = "QC Drop Off"
        
        CommonPath(1) = "\\filecluster01\dfs\3daytemp"
        CommonPath(2) = SupplyNetPATH
        CommonPath(3) = SupplyNetPATH & "\Analytics\DAT Resources"
        CommonPath(4) = "\\filecluster01\dfs\VhaSecure2\SupplyNetworkAnalyticsCorp"
        CommonPath(5) = SupplyNetPATH & "\Analytics\DAT Resources\PSC Cheat Sheets"
        CommonPath(6) = ZeusPATH
        CommonPath(7) = BenchPATH
        CommonPath(8) = TMreviewPATH

        For itm = 1 To UBound(CommonArray) '9
            FoldersCommonList.AddItem CommonArray(itm)
        Next
        
        'Populate templates
        '------------------------
        TemplateArray(1) = "Contract Info"
        TemplateArray(2) = "Xref"
        TemplateArray(3) = "ASF extract"
        TemplateArray(4) = "Report Master"
        TemplateArray(5) = "QC Checklist"
        TemplateArray(6) = "BRD"
        
        For itm = 1 To UBound(TemplateArray)
            FoldersTemplateList.AddItem TemplateArray(itm)
        Next
        
        FolderInitFlg = 1
    
    End If
    
    If Not Trim(PSCVar) = "" Then 'Not ReportInitflg = 1 Then
        
        'populate Report Files (init separately cause files can be created at anytime)
        '---------------------
        Set objFolder = objFSO.GetFolder(ZeusPATH)
        'RptCNT = 0
        'ReDim ReportArray(0 To 0)
        For itm = 0 To FoldersReportList.ListCount - 1 'ListCount
            FoldersReportList.RemoveItem 0
        Next
        For Each ofile In objFolder.Files
            If InStr(LCase(ofile.Name), LCase(FileName_PSC)) And Not Left(ofile.Name, 1) = "~" Then
'                RptCNT = RptCNT + 1
'                ReDim Preserve ReportArray(0 To RptCNT)
'                ReDim Preserve ReportPath(0 To RptCNT)
                LstNm = Replace(ofile.Name, "(" & FileName_PSC & ")", "")
                'LstNm = Replace(ofile.Name, FileName_PSC, "")
'                ReportArray(RptCNT) = Replace(LstNm, ".xlsx", "")
'                ReportPath(RptCNT) = ofile.Path
                FoldersReportList.AddItem Replace(LstNm, ".xlsx", "")
            End If
        Next
        'ReportInitflg = 1
    
    End If
    
    ZeusForm.ZeusPages.Height = FoldersFrame.Top + FoldersFrame.Height + 10
    ZeusForm.ZeusPages.Width = NetInitGo.Left + NetInitGo.Width + GeneralWidthVar
    ZeusForm.Height = ZeusPages.Height + 20
    ZeusForm.Width = ZeusPages.Width
    

'Tools
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 7 Then
    
    ZeusPages.Height = 280 '205 '358
    ZeusPages.Width = 605
    ZeusForm.Height = ZeusPages.Height
    ZeusForm.Width = ZeusPages.Width


'PAT
'=================================================================================================
'=================================================================================================
ElseIf ZeusPages.SelectedItem.Index = 8 Then
    
    ZeusPages.Height = GeneralHeightVar
    ZeusPages.Width = xlslabel.Left + xlslabel.Width + GeneralWidthVar
    ZeusForm.Height = ZeusPages.Height
    ZeusForm.Width = ZeusPages.Width

End If

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlePreReport:
On Error GoTo 0
Resume PreReport

End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PAT/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub ValidationsBttn_Click()

If Not xlsChk.Value = True Then
    Workbooks.Open Validations_xlsxPATH
Else
    Workbooks.Open Validations_xlsPATH
End If

End Sub
Sub ValidationUploadBttn_Click()

Workbooks.Open ValidationUploadPATH

End Sub
Sub MbrBreakoutBttn_Click()

If Not xlsChk.Value = True Then
    Workbooks.Open MbrBreakout_xlsxPATH
Else
    Workbooks.Open MbrBreakout_xlsPATH
End If

End Sub
Sub MbrBreakoutComboBttn_Click()

Workbooks.Open MbrBreakoutComboPATH

End Sub
Sub RFEMacroBttn_Click()

Workbooks.Open RFEMacroPath

End Sub
Sub BRDBttn_Click()

Call BRD

End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Setup/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub AsscXref_Click()

Call setFileCaption("AsscXref")

End Sub
Private Sub AsscExtract_Click()

Call setFileCaption("AsscExtract")

End Sub
Private Sub AsscWF_Click()

Call setFileCaption("AsscWF")

End Sub
Sub setFileCaption(Fname As String)

    ChDir ZeusPATH
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    intchoice = Application.FileDialog(msoFileDialogOpen).Show
1   If intchoice <> 0 Then
        wbstr = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        ZeusForm.Controls(Fname).Caption = " " & Mid(wbstr, InStrRev(wbstr, "\") + 1, Len(wbstr))
    End If

End Sub
Private Sub asscEndDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim recset As New ADODB.Recordset
Dim conn As New ADODB.Connection

If Not Trim(asscEndDate.Text) = "" Then
    If Trim(asscNetwork.Value) = "" Then
        MsgBox ("Please input network first")
        asscEndDate.Text = ""
        AnnualizedChk.SetFocus
    ElseIf Not Len(asscEndDate.Text) > 5 Or Not InStr(asscEndDate.Text, "/") > 0 Then
        MsgBox ("Date must be in mm/yyyy format.  Please re-enter.")
        asscEndDate.Text = ""
        AnnualizedChk.SetFocus
    Else
        'check to see if date is within bounds of current dataset for network
        '--------------------
        'On Error GoTo ERR_NoDates
        sqlstr = "SELECT max(End_Date) FROM Dates "
        recset.Open sqlstr, SpendConn, adOpenStatic, adLockReadOnly
        On Error GoTo 0
        If asscEndDate.Value > recset.Fields(0) Then
            MsgBox ("Date entered is outside the bounds of the dataset.")
            asscEndDate.Text = ""
            AnnualizedChk.SetFocus
        End If
    End If
End If


End Sub
Private Sub asscStartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim recset As New ADODB.Recordset
Dim conn As New ADODB.Connection

If Not Trim(asscStartDate.Text) = "" Then
    If Trim(asscNetwork.Value) = "" Then
        MsgBox ("Please input network first")
        asscStartDate.Text = ""
        AnnualizedChk.SetFocus
    ElseIf Not Len(asscStartDate.Text) > 5 Or Not InStr(asscStartDate.Text, "/") > 0 Then
        MsgBox ("Date must be in mm/yyyy format.  Please re-enter.")
        asscStartDate.Text = ""
        AnnualizedChk.SetFocus
    Else
        'check to see if date is within bounds of current dataset for network
        '--------------------
        'On Error GoTo ERR_NoDates
        sqlstr = "SELECT max(Start_Date) FROM Dates "
        recset.Open sqlstr, SpendConn, adOpenStatic, adLockReadOnly
        On Error GoTo 0
        If asscStartDate.Value < recset.Fields(0) Then
            MsgBox ("Date entered is outside the bounds of the dataset.")
            asscStartDate.Text = ""
            AnnualizedChk.SetFocus
        End If
    End If
End If

End Sub
Private Sub ContractLookup_Click()

Call ExpandLookup("Contract")

End Sub
Private Sub SystemLookup_Click()

Call ExpandLookup("System")

End Sub
Private Sub MemberLookup_Click()

Call ExpandLookup("Member")

End Sub
Sub ExpandLookup(ctrlType As String)

Set Ctrl = ZeusForm.Controls(ctrlType & "LookupFrame")
If Ctrl.Visible = True Then
    ContractLookupFrame.Visible = False
    SystemLookupFrame.Visible = False
    MemberLookupFrame.Visible = False
    ContractLookup.Caption = "Lookup>"
    SystemLookup.Caption = "Lookup>"
    MemberLookup.Caption = "Lookup>"
    ZeusPages.Width = asscContractFrame.Left + asscContractFrame.Width + GeneralWidthVar
Else
    Ctrl.Visible = True
    ZeusForm.Controls(ctrlType & "Lookup").Caption = "Lookup <"
    ZeusPages.Width = Ctrl.Left + Ctrl.Width + GeneralWidthVar
End If
ZeusForm.Width = ZeusPages.Width + 3

End Sub
Private Sub asscContracts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveSingle("asscContracts")

End Sub
Private Sub ContractsReturned_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveSingle("ContractsReturned")

End Sub
Private Sub ContractRemove_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveAll("asscContracts")

End Sub
Private Sub ContractRemove_Click()

Call RemoveSingle("asscContracts")

End Sub
Private Sub asscSystems_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveSingle("asscSystems")

End Sub
Private Sub SystemsReturned_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveSingle("SystemsReturned")

End Sub
Private Sub SystemRemove_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveAll("asscSystems")

End Sub
Private Sub SystemRemove_Click()

Call RemoveSingle("asscSystems")

End Sub
Private Sub asscMembers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveSingle("asscMembers")

End Sub
Private Sub MembersReturned_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveSingle("MembersReturned")

End Sub
Private Sub MemberRemove_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call RemoveAll("asscMembers")

End Sub
Private Sub MemberRemove_Click()

Call RemoveSingle("asscMembers")

End Sub
Private Sub RemoveAll(ctrlNm As String)

    Set Ctrl = ZeusForm.Controls(ctrlNm)
    For i = 0 To Ctrl.ListCount - 1
        Ctrl.RemoveItem 0
    Next
    Ctrl.Value = ""

End Sub
Private Sub RemoveSingle(ctrlNm As String)

Set Ctrl = ZeusForm.Controls(ctrlNm)
For i = Ctrl.ListCount - 1 To 0 Step -1
    If Ctrl.List(i) = Ctrl.Value Then Ctrl.RemoveItem i
Next
If Ctrl.ListCount = 0 Then
    Ctrl.Value = ""
Else
    Ctrl.Value = Ctrl.List(0)
End If

End Sub
Private Sub ContractAdd_Click()

If ContractSelectionChk.Value = True Then
    For Each c In Selection
        If Not Trim(c.Value) = "" Then asscContracts.AddItem Trim(c.Value)
    Next
    asscContracts.Value = asscContracts.List(0)
Else
    asscContracts.AddItem Trim(ContractAddBox.Value)
    asscContracts.Value = Trim(ContractAddBox.Value)
    ContractAddBox.Value = ""
End If


End Sub
Private Sub SystemAdd_Click()

'look up members for system and initialize system array for that system

If SystemSelectionChk.Value = True Then
    For Each c In Selection
        If Not Trim(c.Value) = "" Then asscSystems.AddItem Trim(c.Value)
    Next
Else
    asscSystems.AddItem Trim(SystemAddBox.Value)
    asscSystems.Value = Trim(SystemAddBox.Value)
    SystemAddBox.Value = ""
End If

End Sub
Private Sub MemberAdd_Click()

If MemberSelectionChk.Value = True Then
    For Each c In Selection
        If Not Trim(c.Value) = "" Then asscMembers.AddItem Trim(c.Value)
    Next
Else
    asscMembers.AddItem Trim(MemberAddBox.Value)
    asscMembers.Value = Trim(MemberAddBox.Value)
    MemberAddBox.Value = ""
End If


End Sub
Sub ContractAddBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
asscContracts.AddItem Trim(ContractAddBox.Value)
asscContracts.Value = Trim(ContractAddBox.Value)
ContractAddBox.Value = ""

End Sub
Sub SystemAddBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
asscSystems.AddItem Trim(SystemAddBox.Value)
asscSystems.Value = Trim(SystemAddBox.Value)
SystemAddBox.Value = ""

End Sub
Sub MemberAddBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
asscMembers.AddItem Trim(MemberAddBox.Value)
asscMembers.Value = Trim(MemberAddBox.Value)
MemberAddBox.Value = ""

End Sub
Sub asscPSC_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
Call populateResults("asscContracts", asscPSC.Value)
Call Find_Xref_Extract
asscPSC.SetFocus
'ZeusForm.Caption = asscNetwork.Value & " - " & asscPSC.Value
spendPSC.Value = ""
spendContract.Text = ""
SpendPSCInit = "0"
SpendConInit = "0"


End Sub
Sub ContractPSCcrit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
Call populateResults("ContractsReturned", ContractPSCcrit.Value)

End Sub
Sub SystemNetworkcrit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Dim DBntwk As String
If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
DBntwk = FUN_convDBntwk(SystemNetworkCrit.Value)
Call populateResults("SystemsReturned", DBntwk)

End Sub
Sub MemberSystemcrit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
Call populateResults("MembersReturned", MemberSystemcrit.Value)

End Sub
Sub asscNetwork_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
Call NetworkChange(1)


End Sub
Sub Contractlookupadd_click()

    For i = 0 To ContractsReturned.ListCount - 1
        asscContracts.AddItem ContractsReturned.List(i)
    Next
    asscContracts.Value = asscContracts.List(0)

End Sub
Sub Systemlookupadd_click()

    For i = 0 To SystemsReturned.ListCount - 1
        asscSystems.AddItem SystemsReturned.List(i)
    Next
    asscSystems.Value = asscSystems.List(0)

End Sub
Sub Memberlookupadd_click()

    For i = 0 To MembersReturned.ListCount - 1
        asscMembers.AddItem MembersReturned.List(i)
    Next
    asscMembers.Value = asscMembers.List(0)

End Sub
Private Sub SystemEdit_Click()


'[TBD] have system frame as a visual for adding to multidimensional array, don't add more sysframes
'[TBD] when populating systems from SNA sheet if system not found then check members and see if it's a stand alone member? open a frame for each system and populate with members

'add each member from SNA tab to it's system
'on edit, lookup entire member list for system

'popup member add frame on edit click ***********************

    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset

    'redefine System Arrays
    '-------------------------
    systemnmbr = asscSystems.ListCount
    ReDim Preserve SystemParent(1 To systemnmbr)
    ReDim Preserve SystemNames(0 To systemnmbr)
    For i = 0 To systemnmbr
        SystemNames(i) = asscSystems.List(i)
        If asscSystems.Value = asscSystems.List(i) Then currsystem = i
    Next

'redefine parent array
'-------------------------

ReDim Preserve SystemMembers(1 To UBound(mbrtemparry))
For mbr = 1 To UBound(mbrtemparry)
    SystemMembers(mbr) = mbrtemparry(mbr)
Next
SystemParent(2) = SystemMembers
    
    'find if system frame already exists
    '------------------------------------
    '[TBD] if array does not exist then store system name as it appears in edb in an array and then create a new array with members and populate table
    Set sysFrame = Me.Controls.Add("forms.frame.1", "SystemEditFrame", True)
    sysFrame.Top = SystemEdit.Top
    sysFrame.Left = asscSystems.Left
    
    On Error GoTo errhndlNoArray
    sysheight = 0
    For i = 1 To UBound(Replace(asscSystems.Value, " ", ""))
        Set memLbladd = SysAdd.Controls.Add("Forms.Label.1", Replace(recset.Fields(0), " ", ""), True)
        With DataAdd
            .Caption = 1
            .Font.Size = ArrayModel.Font.Size
            .TextAlign = ArrayModel.TextAlign
            .Width = 170
            .Height = 12
            .Left = 5
            .Top = sysheight
        End With
        Set memChkadd = SysAdd.Controls.Add("Forms.Label.1", Replace(recset.Fields(0), " ", ""), True)
        With DataAdd
            .Caption = 1
            .Font.Size = ArrayModel.Font.Size
            .TextAlign = ArrayModel.TextAlign
            .Width = 170
            .Height = 12
            .Left = 5
            .Top = sysheight
        End With
        sysheight = sysheight + memadd.Height
    Next


    
'    For Each clrframe In Me.Controls
'
'        If TypeName(clrframe) = "Frame" Then
'            If (InStr(clrframe.Name, MSBTarray(2)) Or InStr(clrframe.Name, MSBTarray(3)) Or InStr(clrframe.Name, MSBTarray(4))) And Not InStr(clrframe.Name, "home") Then
'                Me.Controls.Remove clrframe.Name
'            End If
'        ElseIf InStr(clrframe.Name, MSBTarray(1)) Then
'            Me.Controls.Remove clrframe.Name
'        ElseIf InStr(clrframe.Name, "Collapse") Then
'            Me.Controls.Remove clrframe.Name
'        End If
'    Next
'    On Error GoTo 0
'
'    if
'        'if exists then set to visible next to edit button
'        Exit Sub
'    End If
    
    'lookup and populate
    '=============================================================================
NoArray:
    sqlstr = "SELECT DISTINCT Name FROM MEMT1MEINQ WHERE systemname = '" & asscSystems.Value & "' AND Status = 'A' AND not name = ''"
     
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    
    On Error GoTo errhndlNORECSET
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    recset.MoveFirst
    On Error GoTo 0
    
    'ReDim Replace(asscSystems.Value, " ", "")(1 To RecSet.RecordCount)
    
    sysheight = 0
    For i = 1 To recset.RecordCount
        Set memLbladd = SysAdd.Controls.Add("Forms.Label.1", Replace(recset.Fields(0), " ", ""), True)
        With DataAdd
            .Caption = 1
            .Font.Size = ArrayModel.Font.Size
            .TextAlign = ArrayModel.TextAlign
            .Width = 170
            .Height = 12
            .Left = 5
            .Top = sysheight
        End With
        Set memLbladd = SysAdd.Controls.Add("Forms.Label.1", Replace(recset.Fields(0), " ", ""), True)
        With DataAdd
            .Caption = 1
            .Font.Size = ArrayModel.Font.Size
            .TextAlign = ArrayModel.TextAlign
            .Width = 170
            .Height = 12
            .Left = 5
            .Top = sysheight
        End With
        sysheight = sysheight + memadd.Height
        recset.MoveNext
    Next
    
    Set recset = Nothing
    Set conn = Nothing
    
    Application.EnableEvents = True
        

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoArray:
Resume NoArray

'errhndlNoSysFrame:
'ReDim Preserve SystemFrames(1 To SysCNT)
'SystemFrames(SysCNT) = Replace(asscSystems.Value, " ", "")
'
'Resume NoSysFrame

errhndlNORECSET:
Application.EnableEvents = True
Set conn = Nothing
Set recset = Nothing
Exit Sub


End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Folders///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub FoldersPSCList_enter()
    
    If Trim(FoldersNtwkList.Value) = "" Then
        MsgBox "Please select network first"
        Exit Sub
    ElseIf FoldersNtwkList.Value = PrevNetVal Then
        Exit Sub
    End If
    
    PrevNetVal = FoldersNtwkList.Value
    Call SetNetPATH(FoldersNtwkList.Value)
    Call setobjFolder(NetworkPath)
    For itm = 0 To FoldersPSCList.ListCount - 1
        FoldersPSCList.RemoveItem 0
    Next
    For Each initFldr In objFolder.SubFolders
        FoldersPSCList.AddItem initFldr.Name
    Next
   
   
End Sub
Sub NetInitGo_click()

If Trim(FoldersNtwkList.Value) = "" Then
    MsgBox "Please select network first"
    Exit Sub
ElseIf Trim(FoldersPSCList.Value) = "" Then
    initvar = ""
Else
    initvar = FoldersPSCList.Value
End If

Call SetNetPATH(FoldersNtwkList.Value)
retval = Shell("explorer.exe " & NetworkPath & initvar, vbNormalFocus)

'createobject("shell.application").minimizeall  '<--Minimize all windows

End Sub
Sub ToolsGo_click()

For i = 1 To FoldersToolsList.ListCount
    If FoldersToolsList.Value = ToolsArray(i) Then
        If InStr(ToolsPath(i), ".accdb") Then
            AccessToolsApp.OpenCurrentDatabase ToolsPath(i)
            AccessToolsApp.Application.Visible = True
        Else
            On Error Resume Next
            Workbooks.Open (ToolsPath(i))
        End If
    End If
Next


End Sub
Sub CommonGo_click()

For i = 1 To FoldersCommonList.ListCount
    If FoldersCommonList.Value = CommonArray(i) Then retval = Shell("explorer.exe " & CommonPath(i), vbNormalFocus)
Next

End Sub
Sub ReportGo_click()

Workbooks.Open (ZeusPATH & "\" & FoldersReportList.Value & "(" & FileName_PSC & ").xlsx")
'For i = 1 To FoldersReportList.ListCount
'    If FoldersReportList.Value = ReportArray(i) Then Workbooks.Open (ReportPath(i))
'Next

End Sub
Sub TemplateGo_click()

If FoldersTemplateList.Value = TemplateArray(1) Then
    Call ContractInfo_Template
ElseIf FoldersTemplateList.Value = TemplateArray(2) Then
    Call DATxref_Template
ElseIf FoldersTemplateList.Value = TemplateArray(3) Then
    Call Extract_Template
ElseIf FoldersTemplateList.Value = TemplateArray(4) Then
    Call Master_Template
ElseIf FoldersTemplateList.Value = TemplateArray(5) Then
    Call QC_Template
ElseIf FoldersTemplateList.Value = TemplateArray(6) Then
    Call BRD_template
End If


End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Sherlock/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub CleanBttn_Click()

Call METH_ClinicalClean(True)

End Sub
Private Sub SherlockBttn_Click()

SherlockForm.Show (False)

End Sub
Private Sub SherlockISbttn_Click()

Call ClinicalIS

End Sub
Private Sub SherlockOOSbttn_Click()

Call ClinicalOOS

End Sub
Private Sub SherlockTBDbttn_Click()

Call ClinicalTBD

End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Rubiks///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub EinsteinGo_Click()

'Call StrtEinstein
MsgBox ("This button has been disabled at management's behest.")

End Sub
Private Sub HansGo_Click()

RubiksForm.Show (False)

End Sub
Private Sub ConvertUOMsAdd_Click()

CUnmbr = CUnmbr + 1

Set addCUorgnl = ConvertUOMsFrame.Controls.Add("Forms.textbox.1", "CUorgnl" & CUnmbr, True)
With addCUorgnl
    .Height = HomeCUorgnl.Height
    .Width = HomeCUorgnl.Width
    .Left = HomeCUorgnl.Left
    .Top = HomeCUorgnl.Top
End With

Set addCUchg = ConvertUOMsFrame.Controls.Add("Forms.textbox.1", "CUchg" & CUnmbr, True)
With addCUorgnl
    .Height = HomeCUorgnl.Height
    .Width = HomeCUorgnl.Width
    .Left = HomeCUorgnl.Left
    .Top = HomeCUorgnl.Top
End With
Set addCUmfg = ConvertUOMsFrame.Controls.Add("Forms.combobox.1", "CUmfg" & CUnmbr, True)
With addCUmfg
    .Height = HomeCUmfg.Height
    .Width = HomeCUmfg.Width
    .Left = HomeCUmfg.Left
    .Top = HomeCUmfg.Top
End With
Set addCUto = ConvertUOMsFrame.Controls.Add("Forms.label.1", "CUto" & CUnmbr, True)
With addCUto
    .Height = HomeCUto.Height
    .Width = HomeCUto.Width
    .Left = HomeCUto.Left
    .Top = HomeCUto.Top
    .Caption = "to"
    .Font.Bold = HomeCUto.Font.Bold
End With
Set addCUremove = ConvertUOMsFrame.Controls.Add("Forms.commandbutton.1", "CUremove" & CUnmbr, True)
With addCUremove
    .Height = HomeCUremove.Height
    .Width = HomeCUremove.Width
    .Left = HomeCUremove.Left
    .Top = HomeCUremove.Top
    .Caption = "Remove"
End With
ReDim Preserve CUremoves(1 To CUnmbr)
Set CUremoves(CUnmbr).CUremoveEvents = addCUremove

'populate listbox with mfgs
'--------------------------
addCUmfg.AddItem = "All Manufacturers"
Range("ZA:ZA").ClearContents
Range(Range("ZA1").End(xlToRight), Range("ZA2").End(xlToRight)).Offset(0, 1).Value = "x"
For Each c In Range(Range("L3"), Range("N3").End(xlDown).Offset(0, -2))
    If Not c.Value = 0 And Not Trim(c.Value) = "" And Not Application.CountIf(Range(Range("ZA3"), Range("ZA1").End(xlDown)), c.Value) > 0 Then
        Range("ZA1").End(xlDown).Offset(1, 0).Value = c.Value
        addCUmfg.AddItem = c.Value
    End If
Next

ConvertUOMsFrame.Height = ConvertUOMsFrame.Height + addCUto.Height
SandSFrame.Top = SandSFrame.Top + addCUto.Height


End Sub
Sub SpecRngBttn_Click()

If SpecRngBox.Visible = True Then
    SpecRngBox.Visible = False
Else
    SpecRngBox.Visible = True
    On Error GoTo errhndlNOHCO
    Sheets("Line Item Data").Select
    Set temprng = Application.InputBox("Please select range in column N:", Type:=8)
    If temprng.Row = Range("N1").Row Or temprng.Row = Range("N2").Row Then
        SpecRngBox.Text = Range(Range("N3"), Range("N" & temprng.Row + temprng.Rows.Count - 1)).Address(0, 0)
    Else
        SpecRngBox.Text = Range(Range("N" & temprng.Row), Range("N" & temprng.Row + temprng.Rows.Count - 1)).Address(0, 0)
    End If
End If

ZeusPages.Height = GeneralHeightVar
ZeusForm.Height = ZeusPages.Height

Exit Sub:
'::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNOHCO:
Exit Sub

End Sub
Private Sub SpecRngBox_Click()

SpecRngBox.Text = Application.InputBox("Please select range:", Type:=8)

End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TierMax/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub HermesBttn_Click()

Call Working_File_Main

End Sub
Sub IndivGo_click()

    'Setup
    '---------------------------------------------------------
    If Not FUN_Save = vbYes Then Exit Sub
    SetupSwitch = FUN_SetupSwitch
    
    'Select METH
    '--------------------------
    If IndivSelect.Value = IndivTitleArray(1) Then
        Call Import_TierInfo
    ElseIf IndivSelect.Value = IndivTitleArray(2) Then
        Call Import_Pricefile(1)
    ElseIf IndivSelect.Value = IndivTitleArray(3) Then
        Call Import_Pricefile
    ElseIf IndivSelect.Value = IndivTitleArray(4) Then
        Call Refresh_Suppliers
    ElseIf IndivSelect.Value = IndivTitleArray(5) Then
        Call Import_CheatSheet
    ElseIf IndivSelect.Value = IndivTitleArray(6) Then
        Call Import_StdznIndex
    ElseIf IndivSelect.Value = IndivTitleArray(7) Then
        Call Import_Scopeguide
    ElseIf IndivSelect.Value = IndivTitleArray(8) Then
        Call Import_UNSPSC
    ElseIf IndivSelect.Value = IndivTitleArray(9) Then
        Call Import_PRS
    ElseIf IndivSelect.Value = IndivTitleArray(10) Then
        Call Import_PRS(True)
    ElseIf IndivSelect.Value = IndivTitleArray(11) Then
        Call Import_AdminFees
    ElseIf IndivSelect.Value = IndivTitleArray(12) Then
        Call Import_Benchmarking
    ElseIf IndivSelect.Value = IndivTitleArray(13) Then
        Call ExtractCore
        Call FormatCrossRefTabs
    ElseIf IndivSelect.Value = IndivTitleArray(14) Then
        'Call METH_ExtractIntelli
    ElseIf IndivSelect.Value = IndivTitleArray(15) Then
        Call FormatCrossRefTabs
    ElseIf IndivSelect.Value = IndivTitleArray(16) Then
        Call RefreshMembers(3)
    ElseIf IndivSelect.Value = IndivTitleArray(17) Then
        Call StandardizeMfg
    ElseIf IndivSelect.Value = IndivTitleArray(18) Then
        Call METH_NovaPlus
    ElseIf IndivSelect.Value = IndivTitleArray(19) Then
        Call AddSuppliers
        Sheets("initiative spend overview").Range(MSGraphBKMRK, MSGraphBKMRK.End(xlToRight).Offset(MbrNMBR + 1, 0)).Calculate
        'Sheets("initiative spend overview").ChartObjects(2).SeriesCollection(1).DataLabels.Position = xlLabelPositionBestFit
    ElseIf IndivSelect.Value = IndivTitleArray(20) Then
        Call METH_Finalize
    ElseIf IndivSelect.Value = IndivTitleArray(21) Then
        Call Create_Xref
'    ElseIf IndivSelect.Value = IndivTitleArray(21) Then
'        Call EXTRACT_Initialize
'    ElseIf IndivSelect.Value = IndivTitleArray(22) Then
'        Call EXTRACT_Finalize
    ElseIf IndivSelect.Value = IndivTitleArray(22) Then
        Call KeywordGenerator
    ElseIf IndivSelect.Value = IndivTitleArray(23) Then
'        connmbr = Trim(Application.InputBox(prompt:="Add scenario for contract number:", Type:=2))
'        If Not Trim(connmbr) = "" Then Call Add_Scenario(connmbr)
        Call Import_Dates
    End If

    interfaceFlg = 0
    Call FUN_CalcBackOn

End Sub
Private Sub IndivSelect_Change()

'determine Desc
'--------------------------
For i = 1 To UBound(IndivTitleArray)
    If IndivSelect.Value = IndivTitleArray(i) Then
        DescWindow.Caption = IndivDescArray(i)
        Exit For
    End If
Next


End Sub
Private Sub RunZeusBttn_Click()

Call Create_Report
Call FUN_CalcBackOn

End Sub
Sub QCgo_Click()

'Make sure setup for QC
'---------------------------
If Not FUN_Save = vbYes Then Exit Sub
Call RefreshMembers(1)
SetupSwitch = FUN_SetupSwitch(1)

'Run QC
'---------------------------
If QCReviewSelect.Value = "QC" Then
    ReviewFlg = 0
    QCform.Show (False)
    QCform.ScrollTop = 0
    Call QC_Main
ElseIf QCReviewSelect.Value = "Review" Then
    ReviewFlg = 1
    Call QCReview
Else
    MsgBox "Please select QC or Review"
End If


End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Import Spend//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub asscSpendReport_click()

Call setFileCaption("AsscSpendReport")

End Sub
Sub spendSearch_click()

Call Spend_Search

End Sub
Sub spendDiscard_Click()

On Error Resume Next
Range(Sheets("Spend search").Range("A2"), Sheets("Spend search").Range("A" & FUN_lastrow("A")).Offset(1, 0)).EntireRow.Clear
Sheets("Spend search").Cells.Interior.ColorIndex = 0
Sheets("Spend search").Cells.Borders.LineStyle = xlNone

End Sub
Private Sub ImportSpend_Click()

If Not FUN_Save = vbYes Then Exit Sub
SetupSwitch = FUN_SetupSwitch

MainCall = 1
If Trim(ZeusForm.AsscExtract.Caption) = "" Then
    On Error GoTo wrongwb
    Sheets("Spend search").Select
    If Application.CountA(Range("A:A")) > 0 Then
    On Error GoTo 0
    Call StdzExtract    '>>>>>>>>>>
    Call Import_Spend   '>>>>>>>>>>
    Else
        MsgBox ("No spend found in Spend Search tab to import.")
        Call FUN_CalcBackOn
    End If
Else
    Call FUN_TestForSheet("Spend Search")
    Cells.Clear
    wbnm = ActiveWorkbook.Name
    Application.DisplayAlerts = False
    Workbooks.Open (ZeusPATH & Trim(ZeusForm.AsscExtract.Caption))
    Range("A:AQ, BH:BH").Copy
    Workbooks(wbnm).Sheets("Spend Search").Range("A1").PasteSpecial xlPasteAll
    ActiveWorkbook.Close (False)
    Call StdzExtract(1)     '>>>>>>>>>>
    Call Import_Spend    '>>>>>>>>>>
End If
MainCall = 0

Sheets("Spend search").Cells.Clear
Application.StatusBar = False
Call FUN_CalcBackOn

Exit Sub
':::::::::::::::::::::::::::::::::::
wrongwb:
MsgBox "Could not find ""Spend Search"" tab to import from."
Exit Sub

End Sub
Sub NRsrch_change()

    ZeusForm.extractSrch = False
    If Not Trim(ZeusForm.asscNetwork.Text) = "" Then
        If ZeusForm.NRSrch.Value = True Then
            NtwkSource = "NR"
        Else
            NtwkSource = "RDM"
        End If
        Call Connect_To_Dataset
    End If


End Sub
Sub extractSrch_change()

    ZeusForm.NRSrch = False
    If Not Trim(ZeusForm.asscNetwork.Text) = "" Then
        If ZeusForm.extractSrch.Value = True Then
            NtwkSource = "Extract"
        Else
            NtwkSource = "RDM"
        End If
        Call Connect_To_Dataset
    End If


End Sub
Sub BulkReportBttn_Click()

NetworkSelection.Show (False)


End Sub
Sub Find_Xref_Extract()

        FileName_PSC = Trim(Replace(ZeusForm.asscPSC.Value, "/", " "))
        
        'Find Both
        '----------------------------
        If Trim(AsscXref.Caption) = "" And Trim(AsscExtract.Caption) = "" Then
            setobjFolder (ZeusPATH)
            For Each ofile In objFolder.Files
                If InStr(LCase(ofile.Name), "corexref") > 0 And InStr(LCase(ofile.Name), LCase(FileName_PSC)) > 0 And Not InStr(LCase(ofile.Name), "$") > 0 Then
                    AsscXref.Caption = " " & ofile.Name
                ElseIf InStr(LCase(ofile.Name), "asf extract") > 0 And InStr(LCase(ofile.Name), LCase(FileName_PSC)) > 0 And Not InStr(LCase(ofile.Name), "$") > 0 Then
                    AsscExtract.Caption = " " & ofile.Name
                End If
            Next
    
        'Find Xref
        '----------------------------
        ElseIf Trim(AsscXref.Caption) = "" Then
            setobjFolder (ZeusPATH)
            For Each ofile In objFolder.Files
                If InStr(LCase(ofile.Name), "corexref") > 0 And InStr(LCase(ofile.Name), LCase(FileName_PSC)) > 0 And Not InStr(LCase(ofile.Name), "$") > 0 Then
                    AsscXref.Caption = " " & ofile.Name
                    Exit For
                End If
            Next
        
        
        'Find Extract
        '----------------------------
        ElseIf Trim(AsscExtract.Caption) = "" And extractSrch.Value = True Then
            setobjFolder (ZeusPATH)
            For Each ofile In objFolder.Files
                If InStr(LCase(ofile.Name), "asf extract") > 0 And InStr(LCase(ofile.Name), LCase(FileName_PSC)) > 0 Then
                    AsscExtract.Caption = " " & ofile.Name
                    Exit For
                End If
            Next
        End If
        
End Sub

