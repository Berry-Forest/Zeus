Attribute VB_Name = "I__Common"

'DB Tables
'****************************************************
'Server = dwprod.corp.vha.ad
'===============================
'Database = EDB
'Tables:
'--------------------
'OCSDW_CONTRACT
'OCSDW_VENDOR
'OCSDW_PRODUCT
'OCSDW_PROGRAM
'OCSDW_PRODUCT_PACKAGE
'OCSDW_PRICE_Novation
'OCSDW_PRICE_TIER
'OCSDW_CONTRACT_PROGRAM_DETAIL
'OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL
'PRS_MEMBER_CONTRACT_SALES_ALL
'OCSDW_LINE_ITEM_PRICING_DETAIL
'NFMA_Product
'MEMT1MEINQ
'OCSDW_GPO_CONTRACT

'Server = dbvhadmprod
'===============================
'Database = VUN
'Tables:
'--------------------
'plx_bench_stg
'product_master

'Database = EDB
'Tables:
'--------------------
'ocsdw_product
'ocsdw_price

'Server = exap-scan.corp.vha.com (RDM Spend)
'===============================
'Database = toolsprd_rac

'Tables:
'--------------------
'RDM_MBR_R12_MM_SPEND_SVNGS_FCT
'RDM_MBR_AGG_GRP_DIM
'RDM_MBR_DIM
'RDM_MBR_CMPY_DIM
'RDM_SYS_DIM
'RDM_HCO_ITEM_MSTR_DIM
'RDM_NOV_VNDR_MSTR_DIM
'RDM_PRDCT_DIM
'RDM_PRDCT_SPEND_CATGY_DIM
'RDM_UNSPSC_HIERCHY_DIM
'RDM_CONTR_DIM




'Color Theme
'****************************************************
'interior border = 14277081
'outside border = 12566463
'teal font = 11250945 (dark teal, Accent 4)
'orange font = 20223 (orange, Accent 1)
'orange cell bg = 14021375 (Gold, Accent 2, lighter 80%)
'red cell bg = 13425663 (orange, Accent 1, lighter 80%)
'purple cell bg = 15722467

'QC purple = 16711935
'yellow = 65535
'green = 62580


Public PSCArray() As New BulkEventHandler
Public AddConArray() As New BulkEventHandler
Public RmvConArray() As New BulkEventHandler
Public RmvAConArray() As New BulkEventHandler
Public AddConBxArray() As New BulkEventHandler
Public RmvReport() As New BulkEventHandler

'Paths
'---------
Public Const EnvironConfigPath = "\\filecluster01\dfs\NovSecure2\SupplyNetworks\Zeus\DBA\Users\UserAssignment.txt"
Public Const FunctionPath = "\\filecluster01\dfs\NovSecure2\SupplyNetworks\Zeus\Prod\Prod1\Admin\Functions"
Public Usr As String
Public EnvironPath As String
Public ZeusPATH As String
Public TemplatePATH As String
Public uomDirPATH As String
Public NetworkPath As String
Public FileSaveAsName As String     'for saving from IE dialog box with Sub SaveAsURLDialog
Public DrivePATH As String
Public StdMfgPATH As String
Public SNAengagePATH As String
Public SNAengagePATH_DB As String
Public BenchPATH As String
Public XrefdbPATH As String
Public TMreviewPATH As String
Public TMclinicalPATH As String
Public MSreviewPATH As String
Public MSclinicalPATH As String
Public ContractDirPATH As String
Public SupplyNetPATH As String
Public DATResourcePATH As String
Public ScopeguidePATH As String
Public MbrBreakoutComboPATH As String
Public ValidationUploadPATH As String
Public Validations_xlsxPATH As String
Public Validations_xlsPATH As String
Public MbrBreakout_xlsxPATH As String
Public MbrBreakout_xlsPATH As String
Public RFEMacroPath As String
Public NtwkNmArray() As String
Public NtwkIDArray() As Variant
Public NtwkFldrArray() As String
Public NtwkEDBArray() As String
Public MasterTemplate As String
Public BRDTemplate As String
Public LocalComponentsPATH As String
Public LocalDataPATH As String
Public NetworkDataPATH As String
Public UsrEnviron As String
Public AdminConfigPATH As String
Public Stdzn_Index_PATH As String
Public AdminconfigStr As String
Public IconsPATH As String

'Range refs
'---------
Public MbrBkmrk As Range
Public ConTblBKMRK As Range
Public MSGraphBKMRK As Range
Public BenchBKMRK As Range
Public prsBKMRK As Range
Public NonConBKMRK As Range
Public ConvBKMRK As Range
Public LIDSuppBKMRK As Range
Public UOMref
Public ReturnCol
Public ReturnStrt
Public ReturnSet
Public HansSet
Public StdTbl As String
Public RDMTbl As String
Public StdyTbl As String

'Scripting objects
'---------
Public objFolder
Public objFSO
'Public wrdApp As Word.Application
'Public wrdDoc As Word.Document
Public RDMconn As New ADODB.Connection
Public AppAccess2 As New Access.Application
Public AppAccess As Object

'Report varaibles
'---------
Public MbrNMBR As Integer
Public suppNMBR As Integer
Public ItmNmbr As Long
Public PSCVar As String
Public NetNm As String
Public NetPos As Integer
Public Catnmbr As String
Public TMstrtDate As String
Public TMendDate As String
Public SystemFrames() As String
Public SysCNT As Integer
Public CmpyCD As String
Public LargeReport As Boolean
Public FileName_PSC As String

'Workbooks
'---------
Public tmWB As Workbook
Public reportWB As Workbook
Public xrefwb As Workbook
Public bmWB As Workbook
Public CoreXrefWB As Workbook
Public UOMwb As Workbook

'Flags
'---------
Public wbfoundFLG As Integer
Public DoNotOpenFLG As Integer
Public CreateReport As Boolean
Public MainCall As Integer
Public NotReqFLG As Integer
Public endFLG As Integer
Public IndivFlg As Integer
Public CursoryCheckFLG As Integer
Public CalcFLG As Integer
Public UOMfndFLG As Integer
Public msQC As Integer
Public ReviewFlg As Integer
Public HansFLG As Integer
Public PFstdFLG As Integer
Public HansUOMflg As Integer
Public contextFLG As Integer
Public SherlockFLG As Integer
Public ZeusNotes As Integer
Public NotesOpen As Integer
Public SetupSwitch As Integer
Public ShtNotFound As Integer
Public PSCInitFlg As Integer
Public NetInitFlg As Integer
Public ConnNotFnd As Integer
Public QCFlg As Boolean
Public QCChkFlg As Boolean

'Misc
'---------
Public connmbr
Public networkPathVar As String
Public LetterCNT As Integer
Public NumberCNT As Integer
Public ProdNmbr
Public ContextMfg As String
Public HansEAval As Double
Public urlLnk As String
Public WebRefNote As String
Public NetNmbr As Integer
Public AccessToolsApp As New Access.Application
Public SpendConn As New ADODB.Connection
Public CurrRun As String
Public PrevRun As String
Public RDMConnStr As String
Public rdmPwd As String
Public EDBpwd As String
Public AdminUser As String
Public AdminEmail As String
Public NtwkSource As String
Public TeamNm As String
Public SystemNames() As String
Public SystemArray() As Variant
Public MbrToSysNames() As String
Public MbrToSysArray() As Variant
Public MbrNames() As String
Public MbrMIDArray() As Variant
Public OwnrMbrArray() As String
Public OwnrNmbr As Integer
Public ScreenUpdating_YesNo As Boolean

'Userform
'---------
Public FormNM As Object
Public txtCTRL As control
Public lnkCTRL As control
Public Section(1 To 4) As String
Public DataCtrl(1 To 8) As String
Public XoutIcon As String
Public SpendPSCInit As Integer
Public SpendConInit As Integer
Public FUNC_VBS As Variant
Public ReportArray() As Variant

'API
'---------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'for minimize/Maximize
'----------------
Public Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Global Const SW_SHOWMAXIMIZED = 3
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub strtZeusApp(control As IRibbonControl)

ZeusForm.Show (False)

'On Error Resume Next
For Each Wb In Workbooks    '<--Becasue excel 2013 sucks and keeps flipping workbooks
    If InStr(UCase(Wb.Name), UCase(FileName_PSC)) > 0 And InStr(UCase(Wb.Name), UCase(NetNm)) > 0 Then Wb.Activate
'    Workbooks.Application.ActiveWindow.Activate
'Debug.Print Wb.Application.Activate
'Application.Windows(Wb.Name).ActivateNext
Next

If ConnNotFnd = 1 Then
    ZeusForm.Hide
    MsgBox "Zeus could not establish a connection to EDB.  Please check your connection to the server and try again."
    On Error Resume Next
    Unload ZeusForm
    ConnNotFnd = 0
    Application.StatusBar = False
    Exit Sub
End If


End Sub
Sub BulkRun(ReportNtwk, ReportPSC, ConArray)

ZeusForm.Show (False)

ZeusForm.asscNetwork.Text = ReportNtwk
If InStr(FUN_ConvGroups(AdminconfigStr, "CSA Networks"), ReportNtwk) > 0 Then ZeusForm.extractSrch = True
Call NetworkChange(1)

ZeusForm.asscPSC.Text = ReportPSC
For i = 0 To UBound(ConArray)
    If Not Trim(ConArray(i)) = "" Then ZeusForm.asscContracts.AddItem ConArray(i)
Next
ZeusForm.asscContracts.Text = ZeusForm.asscContracts.List(0)
Call FUN_SetupSwitch
Call Create_Report(True)




End Sub
Sub SetPathVariables()


    'Set configuration
    '================================================================================================
    
    'set environconfig
    '---------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set fileObj = objFSO.GetFile(EnvironConfigPath)
    EnvironStr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
    
    'set usr & environ
    '---------------------
    Usr = FUN_findUsr
    UsrEnviron = UCase(Environ("USERPROFILE"))
    EnvironNmbr = FUN_ConvTags(EnvironStr, "Number of Environments")
    For i = 1 To EnvironNmbr
        EnvironUsers = FUN_ConvTags(EnvironStr, "Environ " & i & " Users")
        If InStr(UCase(EnvironUsers), UCase(Usr)) > 0 Then
            ConfigNmbr = 1
            If i = 2 Then
                DEV_YesNo = MsgBox("Would you like to enter DEV environment?", vbYesNo)
                If DEV_YesNo = vbYes Then ConfigNmbr = 2 Else ConfigNmbr = 1
            ElseIf i = 3 Then
                UAT_YesNo = MsgBox("Would you like to enter UAT environment?", vbYesNo)
                If UAT_YesNo = vbYes Then ConfigNmbr = 3 Else ConfigNmbr = 1
            End If
            AdminConfigPATH = FUN_ConvTags(EnvironStr, "Environ " & ConfigNmbr & " Config")
            EnvironPath = FUN_ConvTags(EnvironStr, "Environ " & ConfigNmbr & " Path")
            AdminEmail = FUN_ConvTags(EnvironStr, "Admin Email")
            Exit For
        End If
    Next
    
    'get adminconfig
    '---------------------
    Set fileObj = objFSO.GetFile(AdminConfigPATH)
    AdminconfigStr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
    
    'local paths
    '================================================================================================
    ZeusPATH = FUN_ConvTags(AdminconfigStr, "Local App Folder Path")
    LocalComponentsPATH = FUN_ConvTags(AdminconfigStr, "Local Components Folder Path")
    LocalDataPATH = FUN_ConvTags(AdminconfigStr, "Local Data Folder Path")
    NetworkDataPATH = FUN_ConvTags(AdminconfigStr, "Network Data Path")
'    RDMDataStr = FUN_ConvGroups(AdminconfigStr, "RDM")
'        LocalRDMDataPATH = FUN_ConvTags(RDMDataStr, "Local Folder")
'        NetworkRDMDataPATH = FUN_ConvTags(RDMDataStr, "Network Folder")
'    NRDataStr = FUN_ConvGroups(AdminconfigStr, "NR")
'        LocalNRDataPATH = FUN_ConvTags(NRDataStr, "Local Folder")
'        NetworkNRDataPATH = FUN_ConvTags(NRDataStr, "Network Folder")
'    ExtractDataStr = FUN_ConvGroups(AdminconfigStr, "Extract")
'        LocalExtractDataPATH = FUN_ConvTags(ExtractDataStr, "Local Folder")
'        NetworkExtractDataPATH = FUN_ConvTags(ExtractDataStr, "Network Folder")
    TemplateStr = FUN_ConvGroups(AdminconfigStr, "Templates")
        TemplatePATH = FUN_ConvTags(TemplateStr, "Local Folder")
    MasterTemplate = FUN_ConvTags(AdminconfigStr, "Master Report Template")
    BRDTemplate = FUN_ConvTags(AdminconfigStr, "BRD Template")
    
    'Drive Mapping
    '====================================================================================================
    'DrivePATH = FUN_findDrive
'    If LCase(Usr) = "tingram" Or LCase(Usr) = "lcassida" Then
'        SupplyNetPATH = "J:\NovSecure2\SupplyNetworks\"
'    ElseIf LCase(Usr) = "jokiri" Then
'        SupplyNetPATH = "Z:\"
'    Else
'        SupplyNetPATH = "I:\NovSecure2\SupplyNetworks\"
'    End If
    
    'network paths
    '====================================================================================================
    SupplyNetPATH = FUN_ConvTags(AdminconfigStr, "SupplyNetPATH")
    DATResourcePATH = FUN_ConvTags(AdminconfigStr, "DATResourcePATH")
    StdMfgPATH = FUN_ConvTags(AdminconfigStr, "StdMfgPATH")
    Stdzn_Index_PATH = FUN_ConvTags(AdminconfigStr, "Stdzn_Index_PATH")
    'SNAengagePATH = FUN_ConvTags(AdminConfigStr, "SNAengagePATH")   'ZeusPATH & "1-Tools\Cache\SNA Engagement Standardization.xlsx"
    'SNAengagePATH_DB = FUN_ConvTags(AdminConfigStr, "SNAengagePATH_DB")
    XrefdbPATH = FUN_ConvTags(AdminconfigStr, "XrefdbPATH")
    ScopeguidePATH = Replace(FUN_ConvTags(AdminconfigStr, "ScopeguidePATH"), "**ZEUSPATH**", ZeusPATH) '"\\filecluster01\dfs\NovSecure2\SupplyNetworks\Analytics\DAT Resources\Scope Guide.accdb"
    BenchPATH = FUN_ConvTags(AdminconfigStr, "BenchPATH")
    TMreviewPATH = FUN_ConvTags(AdminconfigStr, "TMreviewPATH")
    TMclinicalPATH = FUN_ConvTags(AdminconfigStr, "TMclinicalPATH")
    'initiativeDirPATH = FUN_ConvTags(AdminconfigStr, "initiativeDirPATH")
    'pscDirPATH = FUN_ConvTags(AdminconfigStr, "pscDirPATH")
    ContractDirPATH = FUN_ConvTags(AdminconfigStr, "ContractDirPATH")
    uomDirPATH = FUN_ConvTags(AdminconfigStr, "uomDirPATH")
    MbrBreakoutComboPATH = FUN_ConvTags(AdminconfigStr, "MbrBreakoutComboPATH")
    MbrBreakout_xlsPATH = FUN_ConvTags(AdminconfigStr, "MbrBreakout_xlsPATH")
    MbrBreakout_xlsxPATH = FUN_ConvTags(AdminconfigStr, "MbrBreakout_xlsxPATH")
    ValidationUploadPATH = FUN_ConvTags(AdminconfigStr, "ValidationUploadPATH")
    Validations_xlsxPATH = FUN_ConvTags(AdminconfigStr, "Validations_xlsxPATH")
    Validations_xlsPATH = FUN_ConvTags(AdminconfigStr, "Validations_xlsPATH")
    RFEMacroPath = FUN_ConvTags(AdminconfigStr, "RFEMacroPath")
    RDMConnStr = FUN_ConvTags(AdminconfigStr, "RDMConnStr")
    
    Call Set_Index_Cols
    
'    'Update Paths
'    '====================================================================================================
'    For grp = 1 To grpnmbr
'        itmnmbr = FUN_ConvTags(AdminConfigStr, "Group " & grp & " Number of Items")
'        LocalGrpFldr = FUN_ConvTags(AdminConfigStr, "Group " & grp & " Local Folder")
'        For itm = 1 To itmnmbr
'            ItmTag = "Group " & grp & " Item " & itm
'            adminitm = FUN_ConvTags(AdminConfigStr, ItmTag)
'            useritm = FUN_ConvTags(UserConfigStr, ItmTag)
'            If objFSO.FileExists(LocalGrpFldr & "\" & adminitm) Then
'                If Not useritm = "" Then
'                    UserConfigStr = Replace(UserConfigStr, "[" & ItmTag & "]" & useritm & "[/" & ItmTag & "]", "[" & ItmTag & "]" & adminitm & "[/" & ItmTag & "]")
'                Else
'                    UserConfigStr = UserConfigStr & vbCrLf & "[" & ItmTag & "]" & adminitm & "[/" & ItmTag & "]"
'                End If
'            End If
'        Next
'    Next
'    If Not UserConfigStr = "" Then
'        Set objConfigFile = objFSO.OpenTextFile(UserConfigPath & "\" & UserConfig, 2)
'        objConfigFile.WriteLine UserConfigStr
'        objConfigFile.Close
'    End If
    

End Sub
Sub SetNetsArray()

    Set fileObj = objFSO.GetFile(AdminConfigPATH)
    AdminconfigStr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
    On Error GoTo 0
    ntwkstr = FUN_ConvTags(AdminconfigStr, "Networks")
    NmStr = FUN_ConvGroups(ntwkstr, "Names")
    ItmStr = FUN_ConvTags(NmStr, "Items")
    'NmStr = Mid(NmStr, InStr(NmStr, "[Group]") + 8, Len(NmStr))
    TempArray = FUN_ItmsToArray(ItmStr)
    For i = 1 To UBound(TempArray)
        ReDim Preserve NtwkNmArray(1 To i)
        NtwkNmArray(i) = TempArray(i)
    Next
'    ItmStr = ItmStr & "[ "
'    Do Until InStr(ItmStr, "][") = 0
'        'TagData = Mid(ItmStr, InStr(ItmStr, "]") + 1, InStr(ItmStr, "[") - InStr(ItmStr, "]") - 1)
'        TagData = Mid(ItmStr, InStr(ItmStr, "]") + 1, InStr(ItmStr, "[/") - InStr(ItmStr, "]") - 1)
'        ItmStr = Mid(ItmStr, InStr(ItmStr, "][") + 1, Len(ItmStr))
'
'        ArryCnt = ArryCnt + 1
'        ReDim Preserve NtwkNmArray(1 To ArryCnt)
'        NtwkNmArray(ArryCnt) = TagData
'    Loop
    NetNmbr = UBound(NtwkNmArray)

    On Error Resume Next
    For i = 1 To NetNmbr
        ReDim Preserve NtwkIDArray(1 To i)
        NtwkIDArray(i) = FUN_ConvTags(ntwkstr, NtwkNmArray(i) & " ID")
        ReDim Preserve NtwkFldrArray(1 To i)
        NtwkFldrArray(i) = FUN_ConvTags(ntwkstr, NtwkNmArray(i) & " Folder")
        ReDim Preserve NtwkEDBArray(1 To i)
        NtwkEDBArray(i) = FUN_ConvTags(ntwkstr, NtwkNmArray(i) & " EDB")
    Next
    
    

End Sub
Sub Set_Index_Cols()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 'NRTbl = "NR"
 'ExtractTbl = "Extract"
 RDMTbl = "RDM"
 StdTbl = "Standardization"
'Set CompTbl = Range("W:Z")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'On Error Resume Next
'
''Study Cols
''-----------------
'StdyStdMIDs_Col = StdyTbl.Find(what:="Std_MID").Column
'ADSMIDs_Col = StdyTbl.Find(what:="ADS_MID").Column
'StdyKeys_Col = StdyTbl.Find(what:="Study_Key").Column
'StdyNms_Col = StdyTbl.Find(what:="Study_Name").Column
'StdySpnd_Col = StdyTbl.Find(what:="Study_Spend").Column
'StdyDts_Col = StdyTbl.Find(what:="Study_Date_Range").Column
'
''RDM Cols
''-----------------
'RDMStdMIDs_Col = RDMTbl.Find(what:="Std_MID").Column
'RDMMIDs_Col = RDMTbl.Find(what:="RDM_MID").Column
'RDMmbrs_Col = RDMTbl.Find(what:="Member_Name").Column
'RDMSIDs_Col = RDMTbl.Find(what:="System_ID").Column
'RDMSys_Col = RDMTbl.Find(what:="System_Name").Column
'RDMSpnd_Col = RDMTbl.Find(what:="RDM_Spend").Column
'RDMDts_Col = RDMTbl.Find(what:="RDM_Date_Range").Column
'
''Std Cols
''-----------------
'StdMIDs_Col = StdTbl.Find(what:="Std_MID").Column
'StdNms_Col = StdTbl.Find(what:="Name_Rolled_Up_To_In_Report").Column
'On Error Resume Next
'StdNms_Rng = Range(Cells(2, StdNms_Col), Cells(Rows.Count, StdNms_Col).End(xlUp).Row)
'StdRDM_Col = StdTbl.Find(what:="RDM_Spend").Column
'StdNR_Col = StdTbl.Find(what:="Study_Spend").Column
'CurrSrce_Col = StdTbl.Find(what:="Current_Source").Column
'FtrSrce_Col = StdTbl.Find(what:="Future_Source").Column
'
''Spend Comparison
''-----------------
'CompMbr_Col = CompTbl.Find(what:="Name_Included_In_Report").Column
'CompNR_Col = CompTbl.Find(what:="Annualized_Network_Run").Column
'CompRDM_Col = CompTbl.Find(what:="12mo_RDM_Extract").Column
'CompVar_Col = CompTbl.Find(what:="Variance").Column


End Sub
Sub Run_VBS_Func()


Set wsh = New WshShell
'Set wsh = CreateObject("WScript.Shell")
'Set FUNC_VBS = wsh.Exec("wscript " & FunctionPath & "\test.vbs 3")
'Set FUNC_VBS = wsh.Exec("wscript \\filecluster01\dfs\NovSecure2\SupplyNetworks\Zeus\Prod\Prod1\Admin\Functions\test.vbs 3")
Set FUNC_VBS = wsh.Exec("wscript C:\Users\Bforrest\Desktop\test.vbs 5")
Do While FUNC_VBS.Status = 0: Loop
Debug.Print "VBS Return Value: "; FUNC_VBS.ExitCode


End Sub
Function FUN_TxtToStr(TxtStr)
    
    Set wsh = New WshShell
    Set FUNC_VBS = wsh.Exec("wscript " & Func_TxtToStr_PATH & " " & TxtStr)
    Do While FUNC_VBS.Status = 0: Loop
    FUN_TxtToStr = FUNC_VBS.ExitCode


End Function
Sub Create_Function_Module()

With ThisWorkbook.VBProject.VBComponents("I__Functions").CodeModule
 .DeleteLines StartLine:=1, Count:=.CountOfLines
 .AddFromFile "C:\Users\Bforrest\Desktop\test.vbs"
 .AddFromFile "C:\Users\Bforrest\Desktop\test.vbs"
End With


End Sub
Function FUN_ConvToStr(TxtStr)

    For i = 1 To Len(TxtStr)
        If Asc(Mid(TxtStr, i, 1)) > 31 And Asc(Mid(TxtStr, i, 1)) < 127 Then FUN_ConvToStr = FUN_ConvToStr & Mid(TxtStr, i, 1)
    Next

End Function
Function FUN_ConvTags(TagStr, TagNm)
    
    Err.Clear
    FUN_ConvTags = Trim(Mid(TagStr, InStr(UCase(TagStr), UCase("[" & TagNm & "]")) + Len(TagNm) + 2, InStr(UCase(TagStr), UCase("[/" & TagNm & "]")) - InStr(UCase(TagStr), UCase("[" & TagNm & "]")) - Len(TagNm) - 2))
    If Err <> 0 Then
        FUN_ConvTags = ""
    Else
        FUN_ConvTags = FUN_ConvVariables(FUN_ConvTags)
    End If
    
End Function
Function FUN_ConvGroups(TagStr, TagNm)
    
    Err.Clear
    FUN_ConvGroups = Trim(Mid(TagStr, InStr(UCase(TagStr), UCase("<" & TagNm & ">")) + Len(TagNm) + 2, Len(TagStr)))
    FUN_ConvGroups = Left(FUN_ConvGroups, InStr(UCase(FUN_ConvGroups), UCase("[/Group]")))
    If Err <> 0 Then
        FUN_ConvGroups = ""
    Else
        FUN_ConvGroups = FUN_ConvVariables(FUN_ConvGroups)
    End If
    
End Function
Function FUN_ConvVariables(VarStr)

    If Err <> 0 Then FUN_ConvVariables = ""
    If InStr(UCase(VarStr), UCase("**USR**")) > 0 Then VarStr = Replace(VarStr, "**USR**", Usr)
    If InStr(UCase(VarStr), UCase("**USRENVIRON**")) > 0 Then VarStr = Replace(VarStr, "**USRENVIRON**", UsrEnviron)
    If InStr(UCase(VarStr), UCase("**ZEUSPATH**")) > 0 Then VarStr = Replace(VarStr, "**ZEUSPATH**", ZeusPATH)
    If InStr(UCase(VarStr), UCase("**LOCALCOMPONENTS**")) > 0 Then VarStr = Replace(VarStr, "**LOCALCOMPONENTS**", LocalComponentsPATH)
    If InStr(UCase(VarStr), UCase("**LOCALDATA**")) > 0 Then VarStr = Replace(VarStr, "**LOCALDATA**", LocalDataPATH)
    If InStr(UCase(VarStr), UCase("**ENVIRONPATH**")) > 0 Then VarStr = Replace(VarStr, "**ENVIRONPATH**", EnvironPath)
    If InStr(UCase(VarStr), UCase("**SUPPLYNETPATH**")) > 0 Then VarStr = Replace(VarStr, "**SUPPLYNETPATH**", SupplyNetPATH)
    If InStr(UCase(VarStr), UCase("**DATRESOURCEPATH**")) > 0 Then VarStr = Replace(VarStr, "**DATRESOURCEPATH**", DATResourcePATH)
    FUN_ConvVariables = VarStr
    
End Function
'Function FUN_ItmsToArray(ttlStr)
'
''put data for each tag in str into an array
'
'    Dim TempArray()
'    Do Until InStr(ttlStr, "[/") = 0
'        ttlStr = Left(ttlStr, InStrRev(ttlStr, "[/") - 2)
'        TagData = Mid(ttlStr, InStrRev(ttlStr, "]") + 1, Len(ttlStr))
'        ArryCnt = ArryCnt + 1
'        ReDim Preserve TempArray(1 To ArryCnt)
'        TempArray(ArryCnt) = TagData
'    Loop
'    FUN_ItmsToArray = TempArray
'
'End Function

Function FUN_ItmsToArray(ttlStr)

    Dim TempArray()
    ttlStr = ttlStr & "[ "
    Do Until InStr(ttlStr, "][") = 0
        'TagData = Mid(ItmStr, InStr(ItmStr, "]") + 1, InStr(ItmStr, "[") - InStr(ItmStr, "]") - 1)
        TagData = Mid(ttlStr, InStr(ttlStr, "]") + 1, InStr(ttlStr, "[/") - InStr(ttlStr, "]") - 1)
        ttlStr = Mid(ttlStr, InStr(ttlStr, "][") + 1, Len(ttlStr))
        
        ArryCnt = ArryCnt + 1
        ReDim Preserve TempArray(1 To ArryCnt)
        TempArray(ArryCnt) = TagData
    Loop
    FUN_ItmsToArray = TempArray
    
    
End Function

Function FUN_findDrive() As String
    
    On Error GoTo errhndlInputDrive
    If Not Dir("I:\", vbDirectory) = vbNullString Then
        FUN_findDrive = "I:\"
    ElseIf Not Dir("J:\", vbDirectory) = vbNullString Then
        FUN_findDrive = "J:\"
    ElseIf Not Dir("Z:\", vbDirectory) = vbNullString Then
        FUN_findDrive = "Z:\"
    ElseIf Not Dir("N:\", vbDirectory) = vbNullString Then
        FUN_findDrive = "N:\"
    End If
    If ZeusInit = 1 Then
        If FUN_findDrive = "" Then FUN_findDrive = Application.InputBox("Please enter the letter of your SNA network drive", Type:=2) & ":\"
    End If
    
Exit Function
':::::::::::::::::::::
errhndlInputDrive:
Resume inputDrive
inputDrive:
On Error Resume Next
If ZeusInit = 1 Then FUN_findDrive = Application.InputBox("Please enter the letter of your SNA network drive", Type:=2) & ":\"
Exit Function


    
End Function
Sub SetNetPATH(NetVar As String)

    'set netpos
    '-----------------------------------------
    For i = 1 To UBound(NtwkNmArray)
        If NtwkNmArray(i) = Trim(NetVar) Then pathVarPos = i
    Next
    
    'set network folder initiatives path
    '-----------------------------------------
    If Trim(NetVar) = "Northeast PC" Then
        NetworkPath = "\\filecluster01\dfs\NovSecure2\SupplyNetworks\CSN " & NtwkFldrArray(pathVarPos) & "\Initiatives\"
    Else
        NetworkPath = "\\filecluster01\dfs\NovSecure2\SupplyNetworks\CSN " & NtwkFldrArray(pathVarPos) & "\"
    End If


End Sub
Function FUN_convDBntwk(NetNm As String)

For i = 1 To UBound(NtwkNmArray)
    If NtwkNmArray(i) = NetNm Then FUN_convDBntwk = NtwkEDBArray(i)
Next

End Function
Sub extractPFnames()

Sheets("Notes").Range("AA:AA").ClearContents
Sheets("Notes").Range("AA1").Value = "SUPPLIER ALIASES"
aliascnt = 0
For Each sht In ActiveWorkbook.Worksheets
    If InStr(LCase(sht.Name), "pricing") > 0 And Not InStr(LCase(sht.Name), "pricing") = 1 Then
        aliascnt = aliascnt + 1
        sht.Visible = True
        Sheets("Notes").Range("AA1").Offset(aliascnt, 0).Value = Left(sht.Name, Len(sht.Name) - 8)
    End If
Next

End Sub
Sub SetBKMRKs()

    Set MbrBkmrk = Sheets("Index").Range("H9")
    Set ConTblBKMRK = Sheets("Index").Range("C:C").Find(what:="Supplier Name")
    Set MSGraphBKMRK = Sheets("Initiative Spend Overview").Range("B10")
    Set BenchBKMRK = Sheets("Initiative Spend Overview").Range("B:B").Find(what:="Benchmarking", lookat:=xlPart).Offset(2, 0)
    Set prsBKMRK = Sheets("Initiative Spend Overview").Range("B:B").Find(what:="PRS", lookat:=xlPart).Offset(2, 0)
    Set NonConBKMRK = Sheets("Vizient Contracts - NC").Range("B11")
    Set ConvBKMRK = Sheets("Vizient Contracts - Conv").Range("B11")
    Set LIDSuppBKMRK = Sheets("Line Item Data").Range("BG4")
    


    

End Sub
Sub SetApplicationStartEvents(events As String)

    On Error Resume Next
    If events = "all" Then
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.CalculateBeforeSave = False
        Application.ScreenUpdating = False
        Application.AutoRecover.Enabled = False
        On Error GoTo 0
        Exit Sub
    End If
    If events = "event" Then Application.EnableEvents = False
    If events = "alerts" Then Application.DisplayAlerts = False
    If events = "calc" Then Application.Calculation = xlCalculationManual
    If events = "calcSave" Then Application.CalculateBeforeSave = False
    If events = "screen" Then Application.ScreenUpdating = False
    If events = "recover" Then Application.AutoRecover.Enabled = False
    On Error GoTo 0

End Sub
Sub setobjFolder(pathNm As String)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(pathNm)

End Sub
Function FUN_AlphaOnly(ConvSTR As String)

Dim c As Characters

FUN_AlphaOnly = ""
For i = 1 To Len(ConvSTR)
    Debug.Print Mid(ConvSTR, i, 1)
    X = Asc(UCase(Mid(ConvSTR, i, 1)))  'Asc(UCase(c))
    If (X > 64 And X < 91) Or X = 32 Then     '32=spaces, Or (x > 47 And x < 58) =numbers
        FUN_AlphaOnly = FUN_AlphaOnly & Mid(ConvSTR, i, 1)
    Else
        FUN_AlphaOnly = FUN_AlphaOnly & " "  '(convert all else to spaces)
    End If
Next

FUN_AlphaOnly = Trim(Replace(FUN_AlphaOnly, "  ", " "))
FUN_AlphaOnly = Trim(Replace(FUN_AlphaOnly, "  ", " "))     '(repeat incase there are 3 to 4 spaces in a row in the middle)
'Debug.Print fun_AlphaOnly

End Function
Function FUN_findUsr() As String

If (UCase(Environ$("Username"))) = "JOWUOR" Then
    FUN_findUsr = "jokiri"
Else
    FUN_findUsr = (Environ$("Username"))
End If

End Function
Function FUN_ColRng(FstRw As Integer, ColIdx As Integer, Optional sht As Worksheet) As Range

On Error Resume Next
'On Error GoTo 0
If sht.Name = "" Then Set sht = ActiveSheet

lstrw = sht.Cells(sht.Rows.Count, ColIdx).End(xlUp).Row

If FstRw = 0 Then FstRw = 2

Set FUN_ColRng = Range(sht.Cells(FstRw, ColIdx), sht.Cells(lstrw, ColIdx))


End Function
Function FUN_lastrow(col As Variant, Optional shtnm As String) As Long

If shtnm = "" Then
    Set sht = ActiveSheet
Else
    Set sht = Sheets(shtnm)
End If

On Error GoTo errhndlTOOMANY
FUN_lastrow = sht.Cells(sht.Rows.Count, col).End(xlUp).Row

Exit Function
'::::::::::::::::::::::::
errhndlTOOMANY:
If IsNumeric(col) Then
    FUN_lastrow = sht.Range("A1").Offset(0, col - 1).End(xlDown).Row
Else
    FUN_lastrow = sht.Range(col & 1).End(xlDown).Row
End If
Resume Next

End Function
'Function FUN_lastrw(col As Integer, Optional sht As String) As Integer
'
'If IsEmpty(sht) Then sht = ActiveSheet.Name
'FUN_lastrow = sht.Cells(Sheets(sht).Rows.Count, col).End(xlUp).Row
'
'End Function
Function FUN_lastcol(Rw As Integer, Optional shtnm As String) As Integer

If shtnm = "" Then
    Set sht = ActiveSheet
Else
    Set sht = Sheets(shtnm)
End If
FUN_lastcol = sht.Cells(Rw, sht.Columns.Count).End(xlToLeft).Column

End Function
Sub FUN_Sort(sheet As String, Rng As Range, critRngOne As Range, adOne As Integer, Optional critRngTwo As Range, Optional adTwo As Integer, Optional critRngThree As Range, Optional adThree As Integer)

'1=ascending

'Clear filter if exists
'-----------------------------
Application.StatusBar = "Sorting"
If Worksheets(sheet).AutoFilterMode = True Then Worksheets(sheet).AutoFilterMode = False
ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Clear

'For 3 levels
'-----------------------------
If Not critRngThree Is Nothing Then
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add Key _
        :=critRngOne, SortOn:=xlSortOnValues, Order:=adOne, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add Key _
        :=critRngTwo, SortOn:=xlSortOnValues, Order:=adTwo, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add Key _
        :=critRngThree, SortOn:=xlSortOnValues, Order:=adThree, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheet).Sort
        .SetRange Rng
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'For 2 levels
'-----------------------------
ElseIf Not critRngTwo Is Nothing Then
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add Key _
        :=critRngOne, SortOn:=xlSortOnValues, Order:=adOne, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add Key _
        :=critRngTwo, SortOn:=xlSortOnValues, Order:=adTwo, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheet).Sort
        .SetRange Rng
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'For 1 level
'-----------------------------
Else
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add Key _
        :=critRngOne, SortOn:=xlSortOnValues, Order:=adOne, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheet).Sort
        .SetRange Rng
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

Application.StatusBar = False

End Sub
Public Function FUN_OpenWBconst(wbPath As String, wbName As String) As String
    
wbfoundFLG = 1
On Error Resume Next
Set openWBchk = Nothing
Set openWBchk = Workbooks(wbName)
If openWBchk Is Nothing Or IsEmpty(openWBchk) Then
    On Error GoTo errhndleNFOUND
    Workbooks.Open (wbPath & wbName)
    FUN_OpenWBconst = wbName
    On Error GoTo 0
Else
    FUN_OpenWBconst = wbName
End If
    
Exit Function
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndleNFOUND:
MsgBox "Please select associated " & wbName & " file."
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intchoice = Application.FileDialog(msoFileDialogOpen).Show
If intchoice <> 0 Then
    wbstr = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Workbooks.Open (wbstr)
    FUN_OpenWBconst = wbName
Else
    wbfoundFLG = 0
    FUN_OpenWBconst = ActiveWorkbook.Name
End If
    
End Function
Public Function FUN_OpenWBvar(srchPath As String, srchName As Variant, srchValTwo As Variant) As String

Application.DisplayAlerts = False
Application.EnableEvents = False
On Error GoTo errhndleNFOUND

wbfoundFLG = 1

'search open workbooks
'-----------------------------------
For Each Wb In Workbooks
    If InStr(LCase(Wb.Name), LCase(srchName)) > 0 And InStr(LCase(Wb.Name), LCase(srchValTwo)) > 0 Then
        FUN_OpenWBvar = Wb.Name
        Exit Function
    End If
Next

'Search Folder
'-----------------------------------
Call setobjFolder(srchPath)
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFolder = objFSO.GetFolder(srchPath)
For Each ofile In objFolder.Files
    Debug.Print ofile.Name
    If InStr(LCase(ofile.Name), LCase(srchName)) > 0 And InStr(LCase(ofile.Name), LCase(srchValTwo)) > 0 Then
        If Not DoNotOpenFLG = 1 Then Workbooks.Open (srchPath & "/" & ofile.Name)
        FUN_OpenWBvar = ofile.Name
        Exit Function
    End If
Next

If FUN_OpenWBvar = "" Or IsEmpty(FUN_OpenWBvar) Then
    GoTo errhndleNFOUND
End If

On Error GoTo 0
WBvarFUNexit: Exit Function
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndleNFOUND:
Resume manualFind
manualFind:
    If NotReqFLG = 1 Then
        MsgBox "Please select associated " & Replace(LCase(srchName), "file", "") & " file.  If no " & Replace(LCase(srchName), "file", "") & " file exists please press cancel."
    ElseIf CursoryCheckFLG = 1 Then
        intchoice = 0
        GoTo 1
    Else
        MsgBox "Please select associated " & Replace(LCase(srchName), "file", "") & " file."
    End If
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    intchoice = Application.FileDialog(msoFileDialogOpen).Show
1   If intchoice <> 0 Then
        wbstr = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        If Not DoNotOpenFLG = 1 Then Workbooks.Open (wbstr)
        FUN_OpenWBvar = Mid(wbstr, InStrRev(wbstr, "\") + 1, Len(wbstr))
    Else
        wbfoundFLG = 0
        On Error GoTo WBvarFUNexit
        FUN_OpenWBvar = ActiveWorkbook.Name
    End If
    
End Function
Sub ScreenUpdating_Off()

    ScreenUpdating_YesNo = Application.ScreenUpdating
    Application.ScreenUpdating = False

End Sub
Sub ScreenUpdating_Resume()

    Application.ScreenUpdating = ScreenUpdating_YesNo

End Sub
Sub FUN_CalcOff()

    'Turn calculations off
    '-------------------------
    On Error GoTo errhndlNoWB
    ActiveWorkbook.Activate
    If Application.Calculation = xlCalculationManual Then
        CalcFLG = 1
    Else
        CalcFLG = 0
    End If
    Application.Calculation = xlCalculationManual

Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoWB:
Workbooks.Add
Resume Next
    
End Sub
Sub FUN_CalcBackOn()

'Turn calc back on
'***************************

If CalcFLG = 1 Then
    Application.Calculation = xlCalculationManual
Else
    'Application.StatusBar = "Excel is caluclating...Please Wait"
    Application.Calculation = xlCalculationAutomatic
    'Application.StatusBar = False
    CalcFLG = 0
End If

End Sub
Function FUN_convAlphaNumeric(parsStr As String) As String

Dim c As Characters

retval = ""
For Each c In parsStr
X = Asc(UCase(c))
If (X > 64 And X < 91) Or (X > 47 And X < 58) Then 'or x=32 (Leave spaces in)
    retval = retval & X
End If

FUN_convAlphaNumeric = retval

End Function
Function FUN_AlphaNumeric(parsStr As String, Optional spaces As Integer) As String

'Dim c As Characters
'Dim retval As String
Dim i As Integer
    
If spaces = 1 Then
    For i = 1 To Len(parsStr)
        X = Asc(UCase(Mid(parsStr, i, 1)))
        If (X > 64 And X < 91) Or (X > 47 And X < 58) Or X = 32 Then '32 (Leave spaces in)
            FUN_AlphaNumeric = FUN_AlphaNumeric & Mid(parsStr, i, 1)
        End If
    Next
Else
    For i = 1 To Len(parsStr)
        X = Asc(UCase(Mid(parsStr, i, 1)))
        If (X > 64 And X < 91) Or (X > 47 And X < 58) Then '32 (Leave spaces in)
            FUN_AlphaNumeric = FUN_AlphaNumeric & Mid(parsStr, i, 1)
        End If
    Next
End If

'FUN_convAlphaNumeric = retval

End Function
Function FUN_convMbr(parsStr As String) As String

Dim i As Byte
    
    For i = 1 To Len(parsStr)
        If Asc(Mid(parsStr, i, 1)) > 31 And Asc(Mid(parsStr, i, 1)) < 127 Then '(x > 64 And x < 91) Or (x > 47 And x < 58) Or x = 32 Then '32 (Leave spaces in)
            FUN_convMbr = FUN_convMbr & Mid(parsStr, i, 1)
        End If
    Next


End Function
Function FUN_convCatnum(parsStr As String) As String

Dim i As Byte
    
    For i = 1 To Len(parsStr)
        If Not Asc(Mid(parsStr, i, 1)) = 91 And Not Asc(Mid(parsStr, i, 1)) = 93 And Not Asc(Mid(parsStr, i, 1)) = 126 Then
            FUN_convCatnum = FUN_convCatnum & Mid(parsStr, i, 1)
        End If
    Next


End Function
Function FUN_convAlpha(parsStr As String) As String

Dim c As Characters

retval = ""
For Each c In parsStr
X = Asc(UCase(c))
If (X > 64 And X < 91) Or X = 32 Then '32 (Leave spaces in)
    retval = retval & X
End If

FUN_convAlpha = retval

End Function
Function FUN_ConvNumeric(s As String) As String

    Dim retval As String
    Dim i As Byte

    retval = ""
    For i = 1 To Len(s)
        If (Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9") Or Mid(s, i, 1) = "/" Or Mid(s, i, 1) = "-" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    FUN_ConvNumeric = retval
    
End Function
Function FUN_NumberOnly(s As String) As String

    Dim retval As String
    Dim i As Byte

    retval = ""
    For i = 1 To Len(s)
        If (Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9") Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    FUN_NumberOnly = retval
    
End Function
Function FUN_LetterNumber(CNTstr As String) As Integer

    Dim cs As Characters
    LetterCNT = 0
    NumberCNT = 0
    
    'For Each cs In CNTstr
        X = Asc(UCase(cs))
        If (X > 64 And X < 91) Then
            NumberCNT = NumberCNT + 1
        ElseIf (X > 47 And X < 58) Then
            LetterCNT = LetterCNT + 1
        End If
    'Next
    
    FUN_LetterNumber = Len(CNTstr)
        
End Function
Function FUN_SQLConvList(str As String, delim As String, Optional delimConv As Integer) As String

'loop through string and parse by specified delimiter
'-------------------------------------
Do Until Not InStr(str, delim) > 0
    parsedstr = parsedstr & Trim(Left(str, InStr(str, delim) - 1)) & delim
    str = Mid(str, InStr(str, delim) + 1, Len(str))
Loop
parsedstr = parsedstr & Trim(Replace(str, delim, ""))   '<--Remove specified delimeter from string

If delimConv = 1 Then
    FUN_SQLConvList = parsedstr                         '<--Leave as is
Else
    parsedstr = Replace(parsedstr, "'", "''")           '<--Change single apostophe to double apostrophe
    FUN_SQLConvList = Replace(parsedstr, delim, "','")  '<--Change specified delimiter in string to comma
End If


End Function
Function FUN_convKeywords() As String

For Each c In Selection
    FUN_convKeywords = c.Value & ";" & FUN_convKeywords
Next
FUN_convKeywords = Left(FUN_convKeywords, Len(FUN_convKeywords) - 1)


End Function



Sub FUN_TestForSheet(shtnm As String)

    On Error GoTo errhndlNOSHEET
    Sheets(shtnm).Visible = True
    Sheets(shtnm).Select
    On Error GoTo 0

Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNOSHEET:
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = shtnm
Exit Sub

End Sub
Sub MailMetrics(TOstr As String, SUBJECTstr As String, BODYstr As Variant, Optional CCstr As String, Optional BCCstr As String, Optional DefaultFormat As Integer)

On Error GoTo ERR_NoSend

If DefaultFormat = 1 Then BODYstr = "<BODY style=font-size:11pt;font-family:Calibri>" & BODYstr & "</BODY>"

    'email
    '------------------------------
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    'On Error Resume Next
    With OutMail
        .To = TOstr
        .CC = CCstr
        .BCC = BCCstr
        .Subject = SUBJECTstr
        .HTMLBody = BODYstr
        .send   'or use .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::
ERR_NoSend:
Exit Sub


End Sub
Function FUN_Save() As String

    FUN_Save = MsgBox("Have you saved your file?", vbYesNo)
    If FUN_Save = vbYes Then
        Call FUN_CalcOff
        hwnd = FindWindow(vbNullString, ZeusForm.Caption)
        apiShowWindow hwnd, SW_SHOWMINIMIZED
    End If

End Function
Sub NetworkChange(ReportStatus As Integer)

    Dim DBntwk As String
    Dim tblNames As ADODB.Recordset
    
    'determine network position in the array
    '------------------------------
    NetNm = ZeusForm.asscNetwork.Value
    For i = 1 To UBound(NtwkNmArray)
        If ZeusForm.asscNetwork.Value = NtwkNmArray(i) Then NetPos = i
    Next
    
    'Get Company Code
    '------------------------------
    If NetNm = "CHA" Then
        ZeusForm.AsscCompany.Text = "050"
    ElseIf NetNm = "SEC" Or NetNm = "Commonwealth" Or NetNm = "GMC" Or NetNm = "MAC" Or NetNm = "UC" Then
        ZeusForm.AsscCompany.Text = "020"
    Else
        ZeusForm.AsscCompany.Text = "001"
    End If
    
    'Define team specific variables
    '------------------------------
    If InStr(FUN_ConvGroups(AdminconfigStr, "SNA Networks"), NetNm) > 0 Then
        Stdzn_Index_PATH = FUN_ConvTags(AdminconfigStr, "Stdzn_Index_PATH") & "\SNA"
        ZeusForm.extractSrch = False
        NtwkSource = "RDM"
        StdyTbl = "NR"
        TeamNm = "SNA"
    ElseIf InStr(FUN_ConvGroups(AdminconfigStr, "CSA Networks"), NetNm) > 0 Then
        Stdzn_Index_PATH = FUN_ConvTags(AdminconfigStr, "Stdzn_Index_PATH") & "\CSA"
        ZeusForm.extractSrch = False
        NtwkSource = "RDM"
        StdyTbl = "Extract"
        TeamNm = "CSA"
    End If
    
    'populate sytem and/or member dropdowns
    '------------------------------
    For i = 0 To ZeusForm.asscSystems.ListCount - 1
        ZeusForm.asscSystems.RemoveItem (0)
    Next
    For i = 0 To ZeusForm.asscMembers.ListCount - 1
        ZeusForm.asscMembers.RemoveItem (0)
    Next
    ZeusForm.asscMembers.Value = ""
    
    'populate system and member LOOKUP dropdowns
    '------------------------------
    DBntwk = FUN_convDBntwk(ZeusForm.asscNetwork.Value)
    Call populateResults("MemberSystemCrit", DBntwk)
    If NoRec = 1 Then GoTo DBConnect    '<--If could not connect to DB when populating lookup dropdowns
    
    If ZeusForm.DSdefaultChk.Value = False Then
        
        'if not data services standardization then transfer db systems from system lookup dropdown
        '------------------------------
        For i = 0 To ZeusForm.MemberSystemcrit.ListCount - 1
            ZeusForm.asscSystems.AddItem ZeusForm.MemberSystemcrit.List(i)
        Next
        ZeusForm.asscSystems.Value = ZeusForm.asscSystems.List(0)
        
    Else
        
        'Import members
        '------------------------------
        On Error GoTo ERR_NoWB
        currsht = ActiveSheet.Name
        On Error GoTo 0
        Application.ScreenUpdating = False
        Call RefreshMembers(ReportStatus)
        On Error GoTo initStdNames
        Call FUN_TestForSheet("xxStdNames")
        Call Setup_StrdNames
        Sheets("xxStdNames").Visible = False
        Sheets(currsht).Select
    
    End If
    
DBConnect:
    On Error Resume Next
    
    'change DB connection
    '---------------------------
    Call Connect_To_Dataset
    
    'shtbkmrk = ActiveSheet.Name
    'Call Spend_Search(1)
    'Sheets(shtbkmrk).Select
    
    ZeusForm.asscPSC.SetFocus
    'ZeusForm.Caption = ZeusForm.asscNetwork.Value & " - " & ZeusForm.asscPSC.Value
    ZeusForm.asscSystems.Value = ZeusForm.asscSystems.List(0)
    ZeusForm.asscMembers.Value = ZeusForm.asscMembers.List(0)


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
initStdNames:
Resume DBConnect

ERR_NoWB:
Workbooks.Add
Resume



End Sub
Function FUN_SetupSwitch(Optional ReportVars As Integer, Optional PSCreq As Integer, Optional NetworkReq As Integer, Optional SuppReq As Integer) As Integer

    'set variables from form
    '================================================================================
    On Error Resume Next
    For Each Wb In Workbooks    '<--Becasue excel 2013 sucks and keeps flipping workbooks
        If InStr(UCase(Wb.Name), UCase(FileName_PSC)) > 0 And InStr(UCase(Wb.Name), UCase(NetNm)) > 0 Then Wb.Activate
    Next
    suppNMBR = Trim(ZeusForm.asscContracts.ListCount)
    PSCVar = Trim(ZeusForm.asscPSC.Text)
    FileName_PSC = Trim(Replace(PSCVar, "/", " "))
    NetNm = Trim(ZeusForm.asscNetwork.Text)
    CmpyCD = Trim(ZeusForm.AsscCompany.Text)
    For i = 1 To UBound(NtwkNmArray)
        If NtwkNmArray(i) = Trim(NetNm) Then NetPos = i
    Next
    On Error GoTo 0

    'set variables from report
    '================================================================================
    On Error GoTo errhndlNoWB
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name = "Line Item Data" Then
            Set tmWB = ActiveWorkbook
            Call SetBKMRKs  '>>>>>>>>>>
            ItmNmbr = Application.CountA(Sheets("Line Item Data").Range("X:X")) - 1
            'Call extractPFnames  '>>>>>>>>>>>
            If ReportVars = 1 Then
                Call PopulateSetupTab
            Else
                MbrNMBR = FUN_MbrNmbr
            End If
            FUN_SetupSwitch = 2
            GoTo popReport
        ElseIf sht.Name = "Contract Info" Then
            FUN_SetupSwitch = 1
            GoTo popReport
        End If
    Next
    
nowb:
    'set variables from manual input
    '================================================================================
    On Error GoTo 0
    If PSCreq = 1 And Trim(PSCVar) = "" Then PSCVar = Application.InputBox(prompt:="Please enter PSC:", Title:="PSC not found", Type:=2)
    If NetworkReq = 1 And Trim(NetNm) = "" Then NetNm = Application.InputBox(prompt:="Please enter network:", Title:="Network not found", Type:=2)
    If SuppReq = 1 And suppNMBR = 0 Then suppNMBR = Application.InputBox(prompt:="Please enter number of suppliers:", Title:="Suppliers not found", Type:=1)
    FUN_SetupSwitch = 3
    
popReport:
    'populate variables in setup
    '================================================================================
    If Not ReportVars = 1 Then Exit Function
    
    'check for network change
    '-----------------------------------------
    If Not Trim(NetNm) = "" And Not ZeusForm.asscNetwork.Text = NetNm Then
        ZeusForm.asscNetwork.Text = NetNm
        Call NetworkChange(2)
    End If
    CmpyCD = Trim(ZeusForm.AsscCompany.Text)
    
    ZeusForm.asscPSC.Text = PSCVar
    
    If Not Trim(suppNMBR) = "" And Not FUN_SetupSwitch = 3 Then
        If FUN_SetupSwitch = 2 Then
            Set asscConRng = ConTblBKMRK
        ElseIf FUN_SetupSwitch = 1 Then
            Set asscConRng = Sheets("Contract info").Range("A2")
        End If
        
        For i = 0 To ZeusForm.asscContracts.ListCount - 1
            ZeusForm.asscContracts.RemoveItem (0)
        Next
        For i = 1 To suppNMBR
            ZeusForm.asscContracts.AddItem asscConRng.Offset(1, i).Value
        Next
        ZeusForm.asscContracts.Value = ZeusForm.asscContracts.List(0)
    End If


Exit Function
':::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoWB:
Resume nowb


End Function
Sub PopulateSetupTab()

    MbrNMBR = FUN_MbrNmbr
    suppNMBR = FUN_suppNmbr
    If Not Sheets("index").Range("C7").Value = "(Network - PSC)" Then
        NetNm = FUN_NetNm
        PSCVar = FUN_PSCvar
        FileName_PSC = Replace(PSCVar, "/", " ")
    End If
     
End Sub
Function FUN_suppNmbr()

    FUN_suppNmbr = Application.CountA(Range(ConTblBKMRK, ConTblBKMRK.End(xlToRight))) - 1

End Function
Function FUN_MbrNmbr()

    FUN_MbrNmbr = Application.CountA(Range(MbrBkmrk, MbrBkmrk.End(xlDown))) - 1
    
End Function
Function FUN_NetNm()

    netstr = Sheets("index").Range("C7").Value
    FUN_NetNm = Left(netstr, InStr(netstr, "-") - 2)
    
End Function
Function FUN_PSCvar()

    pscstr = Sheets("index").Range("C7").Value
    FUN_PSCvar = Mid(pscstr, InStr(pscstr, "-") + 2, Len(pscstr))

    
End Function
Sub Connect_To_Dataset()
    
    
    If Dir(LocalDataPATH & "\" & NtwkSource & "\" & NetNm & ".accdb", vbDirectory) = vbNullString Then
        DataPath = NetworkDataPATH & "\" & NtwkSource & "\" & NetNm & ".accdb"
    Else
        DataPath = LocalDataPATH & "\" & NtwkSource & "\" & NetNm & ".accdb"
    End If
    Set SpendConn = Nothing
    SpendConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DataPath
    'SpendConn.Open "Provider=SQLNCLI11;Server=DBSCDMDEV;Database=Data_Acquisition;Uid=bforrest;pwd=Bethoumyvision1"
    'SpendConn.Open "Driver={SQL Server};Server=DBSCDMDEV;Database=Data_Acquisition;Uid=bforrest;pwd=bethoumyvision1"
    'SpendConn.Open "Provider=sqloledb;Server=DBSCDMDEV;Initial Catalog=Data_Acquisition;Uid=bforrest;pwd=Bethoumyvision1"
    'SpendConn.Open "Provider=SQLNCLI11;Server=DBSCDMDEV;Database=Data_Acquisition;Trusted_Connection=yes"
    
End Sub
Sub populateResults(ReturnCtrl As String, CritVal As String)

    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset

    Set rCtrl = ZeusForm.Controls(ReturnCtrl)
    
    For i = 0 To rCtrl.ListCount - 1
        rCtrl.RemoveItem 0
    Next
    rCtrl.Value = ""
    
    NoRec = 0
    If Trim(CritVal) = "" Then
        NoRec = 1
        Exit Sub
    End If
    
    If ReturnCtrl = "ContractsReturned" Then
        sqlstr = "SELECT DISTINCT CONTRACT_NUMBER FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con INNER JOIN OCSDW_GPO_CONTRACT gpo ON con.CONTRACT_ID = gpo.CONTRACT_ID AND gpo.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "' WHERE con.ATTRIBUTE_VALUE_NAME = '" & CritVal & "' and (con.STATUS_KEY = 'ACTIVE' or con.STATUS_KEY = 'expired') AND NOT UPPER(con.CONTRACT_NAME) LIKE '%NOT BID%' AND con.CONTRACT_NUMBER LIKE '[a-z][a-z][0-9]%[0-9]' ORDER BY CONTRACT_NUMBER"
    ElseIf InStr(rCtrl.Name, "Contract") > 0 Then
        strSELECT = "SELECT con.CONTRACT_NUMBER, con.CONTRACT_EFF_DATE, con.VENDOR_KEY"
        strFROM = " FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con"
        strCMPY = " INNER JOIN OCSDW_GPO_CONTRACT gpo ON con.CONTRACT_ID = gpo.CONTRACT_ID AND gpo.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "'"
        strWHERE = " WHERE con.ATTRIBUTE_VALUE_NAME = '" & ZeusForm.asscPSC.Value & "' AND con.STATUS_KEY IN ('ACTIVE','SIGNED','PENDING') AND con.EXPORT_TYPE_KEY = 'M' ORDER BY VENDOR_KEY, CONTRACT_EFF_DATE DESC"
        sqlstr = strSELECT & strFROM & strCMPY & strWHERE
        'sqlstr = "SELECT CONTRACT gre_NUMBER FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con INNER JOIN OCSDW_GPO_CONTRACT gpo ON con.CONTRACT_ID = gpo.CONTRACT_ID AND gpo.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "' WHERE con.ATTRIBUTE_VALUE_NAME = '" & CritVal & "' and (con.STATUS_KEY = 'ACTIVE' or con.STATUS_KEY = 'Signed') AND NOT UPPER(con.CONTRACT_NAME) LIKE '%NOT BID%' AND con.CONTRACT_NUMBER LIKE '[a-z][a-z][0-9]%[0-9]' ORDER BY CONTRACT_NUMBER"
    ElseIf InStr(rCtrl.Name, "System") > 0 Then
        sqlstr = "SELECT DISTINCT systemname FROM MEMT1MEINQ WHERE networkname = '" & Replace(CritVal, "'", "''") & "'"
    ElseIf InStr(rCtrl.Name, "Member") > 0 Then
        sqlstr = "SELECT DISTINCT Name FROM MEMT1MEINQ WHERE systemname = '" & Replace(CritVal, "'", "''") & "' AND Status = 'A'"
    End If
    
    On Error Resume Next
    Set conn = Nothing
    On Error GoTo 0
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    
    On Error GoTo errhndlNORECSET
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    recset.MoveFirst
    On Error GoTo 0
    
    If InStr(rCtrl.Name, "Contract") > 0 And Not ReturnCtrl = "ContractsReturned" Then
        If recset.RecordCount = 1 Then
            rCtrl.AddItem Trim(recset.Fields(0))
        Else
            Set ConSort = New ADODB.Recordset
            ConSort.Fields.Append "Cons", adVarChar, 20
            ConSort.Open
            ConSort.AddNew
            ConSort.Fields(0) = recset.Fields(0)
            prevVndr = recset.Fields(2)
            For i = 1 To recset.RecordCount - 1
                recset.MoveNext
                If Not recset.Fields(2) = prevVndr Then
                    ConSort.AddNew
                    ConSort.Fields(0) = recset.Fields(0)
                    prevVndr = recset.Fields(2)
                End If
            Next
            ConSort.Sort = "Cons"
            For i = 1 To ConSort.RecordCount
                rCtrl.AddItem Trim(ConSort.Fields(0))
                ConSort.MoveNext
            Next
            Set ConSort = Nothing
        End If
    Else
        For i = 1 To recset.RecordCount
            If Not Trim(recset.Fields(0)) = "" Then rCtrl.AddItem Trim(recset.Fields(0))
            recset.MoveNext
        Next
    End If
    
    rCtrl.Value = rCtrl.List(0)
    
    Set recset = Nothing
    Set conn = Nothing
    
    Application.EnableEvents = True
        

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNORECSET:
Application.EnableEvents = True
Set conn = Nothing
Set recset = Nothing
NoRec = 1
Exit Sub

    
End Sub
Function FUN_CmpyCd(ntwk As String) As String

    If ntwk = "CHA" Then
        FUN_CmpyCd = "050"
    ElseIf ntwk = "SEC" Or ntwk = "Commonwealth" Or ntwk = "GMC" Or ntwk = "MAC" Or ntwk = "UC" Then
        FUN_CmpyCd = "020"
    Else
        FUN_CmpyCd = "001"
    End If
                
End Function

Sub destroyVar()

Set tmWB = Nothing
Set ZeusWB = Nothing
Set xrefwb = Nothing

CreateReport = False



End Sub
Function FUN_SuppName(pfPos) As String

Set suppstrt = LIDSuppBKMRK.Offset(0, (pfPos - 1) * 30)
If Application.CountA(suppstrt.EntireColumn) > 4 Then
    If WorksheetFunction.IsFormula(suppstrt.End(xlDown)) = True Then
        FUN_SuppName = suppstrt.End(xlDown).Formula
    Else
        FUN_SuppName = suppstrt.Offset(1, 0).Formula
    End If
    FUN_SuppName = Mid(FUN_SuppName, InStr(FUN_SuppName, "'") + 1, InStr(FUN_SuppName, "'!") - (InStr(FUN_SuppName, "'") + 1))
    FUN_SuppName = Replace(FUN_SuppName, " Pricing", "")
Else
    FUN_SuppName = suppstrt.Offset(-1, 0).Value
End If



End Function
'Function FUN_SupplierSheet(PFpos, XrefPF) As Integer
'
'    For Each sht In tmWB.Sheets
'        If InStr(sht.Name, XrefPF) Then pfcount = pfcount + 1
'        If pfcount = PFpos Then
'            sht.Visible = True
'            sht.Select
'            FUN_SupplierSheet = 1
'            Exit Function
'        End If
'    Next
'    FUN_SupplierSheet = 0
'
'End Function

Sub Error_Report()

'screenshot each element?
'screenshot code
'save wb
'message user that error has been logged and if it happens again in the next 24hours, it will be looked into


'try to setup in a way so that people can fix their own problems, or Roth can fix them
'create DBA feature to number lines in code module by replacing at beginning of line, same for vb scripts
'sub name
'line number or text
'err.description
'zeus version
'report name
'report network, psc, contracts
'copy of report? to local folder that is picked up and copied to network folder by listener
'dba listener emails me when new error logged in error folder, or when 2 or more of the same error


'set up dev environment
'setup on github


End Sub

