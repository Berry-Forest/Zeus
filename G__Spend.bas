Attribute VB_Name = "G__Spend"

Sub Spend_Search(Optional initConn As Integer)

Dim adoRecSet As New ADODB.Recordset
Dim connDB As New ADODB.Connection
Dim TestConn As New ADODB.Connection
Dim SQLvar As String
Dim DescStr As String

    If NetNm = "" Then
        MsgBox "No network selected.  Please enter a network on the setup tab and press enter."
        Exit Sub
    End If
    
    'If you're not creating an initial report the from selection check box for keywords is checked then convert strings in selected range to single semicolon delimited string
    '------------------------
    If Not CreateReport = True And ZeusForm.KywdFromSelectionChk = True Then
        DescStr = FUN_convKeywords
    Else
        DescStr = ZeusForm.spendDesc.Text
    End If
        
    Call FUN_TestForSheet("Spend Search")
    Application.ScreenUpdating = False
    
    'define prameters from search fields
    '==========================================================================================================================
    
    'psc
    '--------------------------
    If Not Trim(ZeusForm.spendPSC.Value) = "" Then                          '<--Check if criteria is entered in field
        FldNm = "UCASE([PSC])"                                              '<--Convert to Uppercase cause PSC field in RDM is uppercase
        'FldNm = "UPPER([PSC])"
        If ZeusForm.sPscOR = True Then                                      '<--Check if field is "AND" or "OR"
            pscOP = " OR"
        Else
            pscOP = " AND"
        End If
        If ZeusForm.sPscDNC = True Then pscOP = pscOP & " NOT"              '<--Check if field is "Does not contain"
        
        SQLvar = UCase(FUN_SQLConvList(ZeusForm.spendPSC.Value, ";"))       '<--FUN_SQLConvList converts string to SQL syntax (See function for specifics)
        pscPARAM = pscOP & " " & FldNm & " IN ('" & SQLvar & "')"           '<--Combine all string elements together to create holistic SQL parameter string
    Else
        ZeusForm.sPscOR = False                                             '<--If field is empty then clear the AND/OR radials on userform
        ZeusForm.sPscAND = False
    End If
    
    'Contract
    '--------------------------
    If Not Trim(ZeusForm.spendContract.Text) = "" Then
        FldNm = "[CONTRACT_NUMBER]"
        If ZeusForm.sContractOR = True Then
            conOP = " OR"
        Else
            conOP = " AND"
        End If
        If ZeusForm.sContractDNC = True Then conOP = conOP & " NOT"
        
        SQLvar = UCase(FUN_SQLConvList(ZeusForm.spendContract.Text, ";"))
        conParam = conOP & " " & FldNm & " IN ('" & SQLvar & "')"
    Else
        ZeusForm.sContractOR = False
        ZeusForm.sContractAND = False
    End If
    
    'catnum
    '--------------------------
    If Not Trim(ZeusForm.spendCatnum.Text) = "" Then
        FldNm = "UCASE([STANDARD_CATALOG])"
        If ZeusForm.sCatnumOR = True Then
            catnumOP = " OR"
        Else
            catnumOP = " AND"
        End If
        If ZeusForm.sCatnumDNC = True Then catnumOP = catnumOP & " NOT"
        
        SQLvar = UCase(FUN_SQLConvList(ZeusForm.spendCatnum.Text, ";"))
        catnumPARAM = catnumOP & " " & FldNm & " IN ('" & SQLvar & "')"
    Else
        ZeusForm.sCatnumOR = False
        ZeusForm.sCatnumAND = False
    End If
    
    'mftr (wildcard rather than literal search)
    '--------------------------
    If Not Trim(ZeusForm.spendMftr.Text) = "" Then
        FldNm = "UCASE([STANDARD_MFTR])"
        If ZeusForm.sMftrOR = True Then
            mftrOP = " OR"
        Else
            mftrOP = " AND"
        End If
        If ZeusForm.sMftrDNC = True Then mftrOP = mftrOP & " NOT"
        
        SQLvar = UCase(FUN_SQLConvList(ZeusForm.spendMftr.Text, ";", 1))
        SQLvar = Replace(SQLvar, "'", "''")
        mftrPARAM = mftrOP & " (" & FldNm & " LIKE '%" & Replace(SQLvar, ";", "%' OR " & FldNm & " LIKE '%") & "%')"
    Else
        ZeusForm.sMftrOR = False
        ZeusForm.sMftrAND = False
    End If
    
    'desc (wildcard rather than literal search)
    '--------------------------
    If Not Trim(ZeusForm.spendDesc.Text) = "" Or ZeusForm.KywdFromSelectionChk = True Then
        FldNm = "UCASE([HCO_ITEM_DESC])"
        'FldNm = "UPPER([HCO_ITEM_DESC])"
        If ZeusForm.sDescAND = True Then
            descOP = " AND"
        Else
            descOP = " OR"
        End If
        If ZeusForm.sDescDNC = True Then descOP = descOP & " NOT"
        
        SQLvar = UCase(FUN_SQLConvList(DescStr, "+", 1))
        SQLvar = FUN_SQLConvList(SQLvar, ";", 1)
        SQLvar = Replace(SQLvar, "'", "''")                             '<--Change single apostophe to double apostrophe
        SQLvar = Replace(SQLvar, "+", "%' AND " & FldNm & " LIKE '%")   '<--Change Zeus 'or' wildcard to SQL wildcard search syntax
        
        descparam = descOP & " ((" & FldNm & " LIKE '%" & Replace(SQLvar, ";", "%') OR (" & FldNm & " LIKE '%") & "%'))"
    Else
        ZeusForm.sDescOR = False
        ZeusForm.sDescAND = False
    End If
    
    'unspsc
    '--------------------------
    If Not Trim(ZeusForm.spendUNSPSC.Text) = "" Then
        FldNm = "[UNSPSC_CMDTY_CD]"
        If ZeusForm.sUnspscOR = True Then
            UnspscOP = " OR"
        Else
            UnspscOP = " AND"
        End If
        If ZeusForm.sUnspscDNC = True Then UnspscOP = UnspscOP & " NOT"
        
        SQLvar = UCase(FUN_SQLConvList(ZeusForm.spendUNSPSC.Text, ";"))
        unspscPARAM = UnspscOP & " " & FldNm & " IN ('" & SQLvar & "')"
    Else
        ZeusForm.sUnspscOR = False
        ZeusForm.sUnspscAND = False
    End If
    
    'pim
    '--------------------------
    If Not Trim(ZeusForm.spendPIM.Text) = "" Then
        FldNm = "[PIM_KEY]"
        If ZeusForm.sPimOR = True Then
            pimOP = " OR"
        Else
            pimOP = " AND"
        End If
        If ZeusForm.sPimDNC = True Then pimOP = pimOP & " NOT"
        
        SQLvar = UCase(FUN_SQLConvList(ZeusForm.spendPIM.Text, ";"))
        pimPARAM = pimOP & " " & FldNm & " IN ('" & SQLvar & "')"
    Else
        ZeusForm.sPimOR = False
        ZeusForm.sPimAND = False
    End If

    'Combine all parameter strings into one string
    '--------------------------
    searchparam = pscPARAM & mftrPARAM & catnumPARAM & descparam & unspscPARAM & pimPARAM & conParam
    If Not Trim(searchparam) = "" Then wherePARAM = " WHERE (" & Mid(searchparam, 5, Len(searchparam)) & ") "   '<--if there are no search criteria then don't add the WHERE clause

    'Date
    '--------------------------
    If Not ZeusForm.NRSrch = True Then                                                                          '<--If datasource is network run then skip cause NR data doesn't have line item dates
        If Not Trim(ZeusForm.asscStartDate.Value) = "" Or Not Trim(ZeusForm.asscEndDate.Value) = "" Then        '<--Check to see if data criteria were entered
            
            'Get start date
            '--------------------------
            If Trim(ZeusForm.asscStartDate.Value) = "" Then
                strtdt = Year(Date) - 4 & Format(Month(Date), "0#")     '<--if no start date specified then set 4 years back (Date format = yyyymm)
            Else
                strtdt = Year(ZeusForm.asscStartDate.Value) & Format(Month(ZeusForm.asscStartDate.Value), "0#")
            End If
            
            'Get end date
            '--------------------------
            If Trim(ZeusForm.asscEndDate.Value) = "" Then
                enddt = Year(Date) & Format(Month(Date), "0#")          '<--if no end date specified then set as current date (Date format = yyyymm)
            Else
                enddt = Year(ZeusForm.asscEndDate.Value) & Format(Month(ZeusForm.asscEndDate.Value), "0#")
            End If
            DateParam = " AND [DATE_SUBMITTED] BETWEEN '" & strtdt & "' AND '" & enddt + 1 & "' "
        End If
    End If
    
    'Find member and systems if not DS default (DS default distinction is currently a bookmark for future use by other users who do not need to use a standardization index)
    '=====================================================================
    If ZeusForm.DSdefaultChk.Value = False Then
    
'xxxxxxxxxxxxxxxxxxxxxxx
'        'Find systems
'        '----------------------
'        If ZeusForm.asscSystems.ListCount > 0 Then
'            On Error GoTo errhndlsysdone
'            For i = 0 To ZeusForm.asscSystems.ListCount
'                sysPARAM = sysPARAM & "'" & Replace(ZeusForm.asscSystems.List(i), "'", "''") & "', "
'            Next
'    sysdone: sysPARAM = " AND [HOSPITAL NAME] IN (" & Left(sysPARAM, Len(sysPARAM) - 2) & ")"
'        End If
'xxxxxxxxxxxxxxxxxxxxxxx
        
        'Get members from member dropdown list
        '----------------------
        If ZeusForm.asscMembers.ListCount > 0 Then                                                      '<--Check to make sure members are present
            For i = 0 To ZeusForm.asscMembers.ListCount - 1
                mbrPARAM = mbrPARAM & "'" & Replace(ZeusForm.asscMembers.List(i), "'", "''") & "', "    '<--convert to SQL syntax and add to mbr parameter string
            Next
            mbrPARAM = " AND [MEMBER_NAME] IN (" & Left(mbrPARAM, Len(mbrPARAM) - 2) & ")"
        End If
    
    End If
    
'''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
''Pull From Server
''[TBD]
''create proc for formatting NR files to upload to server
''Join NR run tables to RDM tables for missing mbrs
''store stdzn index on server and have proc for downloading info, modifying, submitting
'
'
'    Call Connect_To_Dataset  '>>>>>>>>>>
'
'    If ZeusForm.TransLvlBx = True Then
'        strSELECT = "SELECT [UNIQUE_ID], [CLINICAL_NAME], [PSC], [UNSPSC_COMMODITY], [PIM_KEY], [MEMBER_ID], [MEMBER_NAME], [SYSTEM_NAME], [SYSTEM_ID], [NETWORK], [HCO_ITEM_NBR], [HCO_ITEM_DESC], [VENDOR_NAME], [VENDOR_CATALOG], [HCO_MFTR], [STANDARD_MFTR], [HCO_CATALOG], [STANDARD_CATALOG], [QTY_PER_UOM], [UOM_PKG_DESC], [PRICE_PER_UOM], [EACH_PRICE], [REPORTED_USAGE], [TOTAL_UNITS], [TOTAL_SPEND], [MATCH_TYPE], [CONTRACTED_MFTR], [CONTRACT_NUMBER], [CONTRACT_NAME], [NOV_DESC], [NOV_CATALOG], [NOV_UOM_DESC], [NOV_UOM_PRICE], [NOV_EACH_PRICE], [DATE_SUBMITTED] "
'    Else
'        strSELECT = "SELECT MAX([UNIQUE_ID]) AS UNIQUE_ID, [CLINICAL_NAME], [PSC], [UNSPSC_COMMODITY], [PIM_KEY], [MEMBER_ID], [MEMBER_NAME], [SYSTEM_NAME], [SYSTEM_ID], [NETWORK], [HCO_ITEM_NBR], [HCO_ITEM_DESC], [VENDOR_NAME], [VENDOR_CATALOG], [HCO_MFTR], [STANDARD_MFTR], [HCO_CATALOG], [STANDARD_CATALOG], [QTY_PER_UOM], [UOM_PKG_DESC], [PRICE_PER_UOM], [EACH_PRICE], [REPORTED_USAGE], [TOTAL_UNITS], [TOTAL_SPEND], [MATCH_TYPE], [CONTRACTED_MFTR], [CONTRACT_NUMBER], [CONTRACT_NAME], [NOV_DESC], [NOV_CATALOG], [NOV_UOM_DESC], [NOV_UOM_PRICE], [NOV_EACH_PRICE], MAX([DATE_SUBMITTED]) AS DATE_SUBMITTED "
'        strGROUPBY = " GROUP BY [CLINICAL_NAME], [PSC], [UNSPSC_COMMODITY], [PIM_KEY], [MEMBER_ID], [MEMBER_NAME], [SYSTEM_NAME], [SYSTEM_ID], [NETWORK], [HCO_ITEM_NBR], [HCO_ITEM_DESC], [VENDOR_NAME], [VENDOR_CATALOG], [HCO_MFTR], [STANDARD_MFTR], [HCO_CATALOG], [STANDARD_CATALOG], [QTY_PER_UOM], [UOM_PKG_DESC], [PRICE_PER_UOM], [EACH_PRICE], [REPORTED_USAGE], [TOTAL_UNITS], [TOTAL_SPEND], [MATCH_TYPE], [CONTRACTED_MFTR], [CONTRACT_NUMBER], [CONTRACT_NAME], [NOV_DESC], [NOV_CATALOG], [NOV_UOM_DESC], [NOV_UOM_PRICE], [NOV_EACH_PRICE] "
'    End If
'
'    tblNm = "[Spend_" & NtwkSource & "_" & NetNm & "] rdm"
'    If ZeusForm.NRSrch = True Then strFROM = ", [Study_Key] "             '<--if using netowrk run data then grab study key field to be used in determining member standardization"
'    strFROM = strFROM & " From " & tblNm
'    'UNION ALL(select from RDM where stdzn =RDM and NR where stdzn = NR)
'    'strMBR_JOIN = INNER JOIN [Spend_NR_" & netnm & "] nr ON nr.MEMBER_NAME = std.MEMBER_NAME INNER JOIN [stdzfilepath] std ON RDM. MEMBER_ID"
'    strWHERE = wherePARAM & mbrPARAM & sysPARAM & DateParam
'    sqlstr = strSELECT & strFROM & strWHERE & strGROUPBY
'
'    adoRecSet.Open sqlstr, TestConn, adOpenForwardOnly, adLockReadOnly
'
'''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'Establish DB Connection
    '=====================================================================
    '--Connection to local network dataset is done when network is entered or changed in setup.
    '--If datasource is changed then connection is switched to local network dataset of selected type at time of change.

    On Error GoTo ERR_NoConnect
    'On Error GoTo 0
Reconnect:
    If SpendConn = "" Then                                                          '<--If no dataset connection has been established then connection switches to shared dataset on I: drive.
        Call Connect_To_Dataset  '>>>>>>>>>>
'        If ZeusForm.NRSrch = True Then
'            DataPath = NetworkNRDataPATH & "\" & NetNm & ".accdb"
'        ElseIf ZeusForm.extractSrch = True Then
'            DataPath = NetworkExtractDataPATH & "\" & NetNm & ".accdb"
'        Else
'            DataPath = NetworkRDMDataPATH & "\" & NetNm & ".accdb"
'        End If
'        Application.StatusBar = "Connecting to Database...please wait"
'        SpendConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DataPath  '<--Open connection
    End If

'Assemble query parameters into final string
'=====================================================================
    strSELECT = "SELECT [UNIQUE_ID], [CLINICAL_NAME], [PSC], [UNSPSC_COMMODITY], [PIM_KEY], [MEMBER_ID], [MEMBER_NAME], [SYSTEM_NAME], [SYSTEM_ID], [NETWORK], [HCO_ITEM_NBR], [HCO_ITEM_DESC], [VENDOR_NAME], [VENDOR_CATALOG], [HCO_MFTR], [STANDARD_MFTR], [HCO_CATALOG], [STANDARD_CATALOG], [QTY_PER_UOM], [UOM_PKG_DESC], [PRICE_PER_UOM], [EACH_PRICE], [REPORTED_USAGE], [TOTAL_UNITS], [TOTAL_SPEND], [MATCH_TYPE], [CONTRACTED_MFTR], [CONTRACT_NUMBER], [CONTRACT_NAME], [NOV_DESC], [NOV_CATALOG], [NOV_UOM_DESC], [NOV_UOM_PRICE], [NOV_EACH_PRICE], [DATE_SUBMITTED]"
    If ZeusForm.NRSrch = True Then              '<--if using netowrk run data then grab study key field to be used in determining member standardization
        strFROM = ", [Study_Key] FROM [Spend]"
    Else
        strFROM = " From [SPEND]"
    End If
    strWHERE = wherePARAM & mbrPARAM & sysPARAM & DateParam
    sqlstr = strSELECT & strFROM & strWHERE

ExecuteQuery:
'==========================================================================================================================================================
    On Error GoTo ERR_QueryFail
    adoRecSet.Open sqlstr, SpendConn, adOpenForwardOnly, adLockReadOnly     '<--Execute query and return data to virtual recordset
    On Error GoTo 0
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx


'Format Results
'==========================================================================================================================================================
'==========================================================================================================================================================
    
    'setup
    '=============================================
    If Trim(Range("A1").Value) = "" Then                                '<--If headers not already present then add them
        For i = 0 To adoRecSet.Fields.Count - 1
            Range("A1").Offset(0, i).Value = adoRecSet.Fields(i).Name
        Next
        Call Search_Setup  '>>>>>>>>>>                                  '<--Format search tab if not already formatted
    Else
        On Error Resume Next
        Range("A:A").Find(what:="x", lookat:=xlWhole).EntireRow.Delete Shift:=xlUp          '<--Delete Black bar if exists
        Range("A:A").Find(what:="No Data", lookat:=xlWhole).EntireRow.Delete Shift:=xlUp    '<--Delete "No Data" if exists
        Rows("1:1").AutoFilter = False                                                      '<--Remove filter if tab is filtered
        On Error GoTo 0
    End If
    
    Set lastrw = Range("A" & FUN_lastrow("A"))      '<--set lastrow of data
        
    On Error GoTo errhndlNORECSET
    lastrw.Offset(1, 0).CopyFromRecordset adoRecSet '<--Copy query results from virtual recordset to search tab in workbook
    adoRecSet.Close                                 '<--Close virtual recordset
    On Error GoTo 0

    If Trim(lastrw.Offset(1, 0).Value) = "" Then    '<--If no results were returned then mark as no data in workbook and skip the rest of results formatting
        lastrw.Offset(1, 0).Value = "No Data"
        lastrw.Offset(1, 0).Font.ColorIndex = 3
        GoTo searchEnd
    Else
        'make sure Unique IDs in col A are not stored as text
        '---------------------------
        Range("A:A").TextToColumns Destination:=Range("A:A"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
    End If
        
    'Dedup items & delete unwanted members
    '=============================================
    If Not ZeusForm.NRSrch = True Then
        Set M2Mrng = Range(Sheets("xxStdNames").Range("B1"), Sheets("xxStdNames").Range("B1").End(xlDown))
        Set M2Srng = Range(Sheets("xxStdNames").Range("D1"), Sheets("xxStdNames").Range("D1").End(xlDown))
        Set SYSrng = Range(Sheets("xxStdNames").Range("F1"), Sheets("xxStdNames").Range("F1").End(xlDown))
    End If
        
    If lastrw.Row = 1 Then
        Set deduprng = Range("A1")  'dummy rng
    Else
        Set deduprng = Range(Range("A2"), lastrw)
    End If
        
    On Error Resume Next
    Set HCOdedup = Range(Sheets("xxProgramData").Range("A2"), Sheets("xxProgramData").Range("A1").End(xlDown))
    If IsEmpty(HCOdedup) Then Set HCOdedup = Range(Sheets("Line Item Data").Range("A5"), Sheets("Line Item Data").Range("A4").End(xlDown))
    If IsEmpty(HCOdedup) Then Set HCOdedup = Sheets("Spend Search").Range("A1")   'Dummy rng
    On Error GoTo 0
    
    For Each c In Range(lastrw.Offset(1, 0), lastrw.End(xlDown))
        If Application.CountIf(HCOdedup, c.Value) > 0 Or Application.CountIf(deduprng, c.Value) > 0 Then
            Range("AK" & c.Row).Value = 1
        ElseIf Not ZeusForm.NRSrch = True Then
            If Not (Application.CountIf(M2Mrng, Range("F" & c.Row).Value) > 0 Or Application.CountIf(M2Srng, Range("F" & c.Row).Value) > 0 Or Application.CountIf(SYSrng, Range("I" & c.Row).Value) > 0) Then Range("AK" & c.Row).Value = 1
        End If
    Next
        
    On Error Resume Next
    Range("AK:AK").SpecialCells(xlCellTypeConstants, 1).EntireRow.Delete Shift:=xlUp
    On Error GoTo 0
    If Trim(lastrw.Offset(1, 0).Value) = "" Then
        lastrw.Offset(1, 0).Value = "No Data"
        lastrw.Offset(1, 0).Font.ColorIndex = 3
        GoTo searchEnd
    End If
    
    'format
    '---------------------
'    If ZeusForm.NRsrch = True Then
'        Range("A:A,C:C,D:D,E:E,F:F,H:H,L:L,N:N,W:W,Y:Y,AA:AA,AB:AB").Borders(xlEdgeRight).Weight = xlMedium
'        Range("C:C").Interior.Color = 14277081
'        Range("D:D").Interior.Color = 14277081
'        Range("E:E").Interior.Color = 14277081
'        Range("F:F").Interior.Color = 14281213
'        Range("H:H").Interior.Color = 15853276
'        Range("L:L").Interior.Color = 15523812
'        Range("N:N").Interior.Color = 14336204
'        Range("W:W").Interior.Color = 12379352
'        Range("Y:Y").Interior.Color = 8421504
'        Range("AA:AA").Interior.Color = 14281213
'        Range("AB:AB").Interior.Color = 14277081
'    Else
        Range("A:A,C:C,D:D,E:E,G:G,H:H,L:L,P:P,R:R,Y:Y,AA:AA,AB:AB,AC:AC").Borders(xlEdgeRight).Weight = xlMedium
        Range("C:C").Interior.Color = 14277081
        Range("D:D").Interior.Color = 14277081
        Range("E:E").Interior.Color = 14277081
        Range("G:G").Interior.Color = 14281213
        Range("H:H").Interior.Color = 14281213
        Range("L:L").Interior.Color = 15853276
        Range("P:P").Interior.Color = 15523812
        Range("R:R").Interior.Color = 14336204
        Range("Y:Y").Interior.Color = 12379352
        Range("AB:AB").Interior.Color = 14281213
        Range("AC:AC").Interior.Color = 14277081
'    End If

    'potential spend & rows
    '---------------------------
    If ZeusForm.NRSrch = True Then
        Set spendrng = Range("W2:W" & lastrw.End(xlDown).Row)
    Else
        Set spendrng = Range("Y2:Y" & lastrw.End(xlDown).Row)
    End If
    spendrng.TextToColumns Destination:=spendrng, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    On Error GoTo errhndlNOHCO
    ZeusForm.PotSpend.Caption = "Potential Spend:  " & Format(Sheets("Line Item Data").Range("X1").Value + Application.sum(spendrng), "$#,##0")
    ZeusForm.PotRows.Caption = "Potential Rows:    " & Range(Sheets("Line Item Data").Range("A5"), Sheets("Line Item Data").Range("A4").End(xlDown)).Count + Range(Range("A2"), Range("A1").End(xlDown)).Count

searchEnd:
    'Clean and end
    '----------------------------
    On Error Resume Next
    If lastrw.Row > 1 Then
        lastrw.Offset(1, 0).EntireRow.Insert
        lastrw.Offset(1, 0).Value = "x"
        lastrw.Offset(1, 0).EntireRow.Interior.ColorIndex = 1
    End If
    Range(Range("A1"), Range("AK1").End(xlToRight)).AutoFilter
    Application.ScreenUpdating = True
    lastrw.Offset(1, 0).Select
    Application.StatusBar = False
    Debug.Print "End: " & Now
    
    'Flag if creating report with >10,000 lines
    '----------------------------
    If CreateReport = True Then
        If Range("A1").End(xlDown).Row > 10000 Then
            LargeReport = True
            Range(Range("A5000"), Range("A5000").End(xlDown)).EntireRow.Delete Shift:=xlUp
        End If
    End If
        
    

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'errhndlmbrsdone:
'Resume mbrsdone

ERR_NoConnect:
    Set SpendConn = Nothing
    On Error GoTo ERR_NoRecord
Resume Reconnect
    
ERR_NoRecord:
    MsgBox ("Could not connect to data source, please check your connection and try again.")
    endFLG = 1
    Application.StatusBar = False
Exit Sub

ERR_QueryFail:
    If InStr(Err.Description, "Query is too complex.") > 0 Then
        MsgBox ("Too many search criteria, please limit your keywords and try again.")
    Else
        MsgBox ("Could not execute query, please check your criteria and try again.")
    End If
    endFLG = 1
    Application.StatusBar = False
Exit Sub

errhndlNORECSET:
    lastrw.Offset(1, 0).Value = "No Data"
    lastrw.Offset(1, 0).Font.ColorIndex = 3
Resume searchEnd

errhndlNOHCO:
    ZeusForm.PotSpend.Caption = "Potential Spend:  " & Format(Application.sum(spendrng), "$#,##0")
    ZeusForm.PotRows.Caption = "Potential Rows:    " & Range(Range("A2"), Range("A1").End(xlDown)).Count
Resume searchEnd


End Sub
Sub Search_Setup()

        'setup search tab
        '-----------------------
        Call Setup_StrdNames
        Sheets("Spend Search").Select
        Rows("1:1").HorizontalAlignment = xlCenter
        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
        
'        If ZeusForm.NRsrch = True Then
'            Range("B:B,G:G,I:K,M:M,O:V,X:X,Z:Z,AC:AI").Columns.Hidden = True
'            Range("A:A,C:C,D:D,E:E,F:F,H:H,L:L,N:N,W:W,Y:Y,AA:AA,AB:AB").Borders(xlEdgeRight).Weight = xlMedium
'            Range("A:A").ColumnWidth = 10
'            Range("C:C").ColumnWidth = 30
'            Range("D:D").ColumnWidth = 19
'            Range("E:E").ColumnWidth = 10
'            Range("F:F").ColumnWidth = 30
'            Range("H:H").ColumnWidth = 75
'            Range("L:L").ColumnWidth = 32
'            Range("K1").Value = "MFTR_NAME"
'            Range("L1").Value = "STANDARD_MFTR"
'            Range("N:N").ColumnWidth = 25
'            Range("M1").Value = "MFTR_CATALOG"
'            Range("N1").Value = "STANDARD_CATALOG"
'            Range("W:W").ColumnWidth = 12
'            Range("W:W").NumberFormat = "$#,##0"
'            Range("W1").Value = "SPEND"
'            Range("Y:Y").ColumnWidth = 6.5
'            Range("Y:Y").HorizontalAlignment = xlCenter
'            Range("Y1").Value = "MATCH"
'            Range("AA:AA").ColumnWidth = 10
'            Range("AA1").Value = "CONTRACT"
'            Range("AB:AB").ColumnWidth = 40
'            Range("AB1").Value = "CONTRACT_DESC"
'        Else
            Range("B:B,F:F,I:I,J:J,K:K,M:O,Q:Q,S:X,Z:AA,AD:AJ").Columns.Hidden = True
            Range("A:A").ColumnWidth = 10
            Range("C:C").ColumnWidth = 22
            Range("D:D").ColumnWidth = 19
            Range("E:E").ColumnWidth = 10
            Range("G:G").ColumnWidth = 25
            Range("H:H").ColumnWidth = 20
            Range("L:L").ColumnWidth = 50
            Range("P:P").ColumnWidth = 25
            Range("R:R").ColumnWidth = 21
            Range("Y:Y").ColumnWidth = 12
            Range("Y:Y").NumberFormat = "$#,##0"
            Range("AB:AB").ColumnWidth = 10
            Range("AC:AC").ColumnWidth = 40
        'End If
        
End Sub
Sub Setup_StrdNames()


Call FUN_TestForSheet("xxStdNames")
Cells.Clear
Range(Range("A1"), Range("F2")).Value = "x"
On Error Resume Next

''SNA
''-----------------------------
'If ZeusForm.NRsrch = True Then
'    For i = 1 To UBound(IDtoNM)
'        Range("A3").Offset(IDrw).Value = IDtoNM(i, 1)
'        Range("B3").Offset(IDrw).Value = IDtoNM(i, 2)
'        IDrw = IDrw + 1
'    Next
'    For i = 1 To UBound(NMtoNM)
'        Range("C3").Offset(NMrw).Value = NMtoNM(i, 1)
'        Range("D3").Offset(NMrw).Value = NMtoNM(i, 2)
'        NMrw = NMrw + 1
'    Next
'
''RDM
''-----------------------------
'Else
    
    For i = 1 To UBound(MbrMIDArray)
        Application.StatusBar = "Setting Up Member Standardizaiton: " & i & " of " & UBound(MbrMIDArray)
        For k = 1 To UBound(MbrMIDArray(i))
            Range("A3").Offset(mbrrw).Value = MbrNames(i)
            Range("B3").Offset(mbrrw).Value = MbrMIDArray(i)(k)
            mbrrw = mbrrw + 1
        Next
    Next
    For i = 1 To UBound(MbrToSysArray)
        For k = 1 To UBound(MbrToSysArray(i))
            Range("C3").Offset(SysRw).Value = MbrToSysNames(i)
            Range("D3").Offset(SysRw).Value = MbrToSysArray(i)(k)
            SysRw = SysRw + 1
        Next
    Next
    For i = 1 To UBound(SystemArray)
        Range("E3").Offset(i).Value = SystemNames(1, 1)
        Range("F3").Offset(i).Value = SystemNames(1, 2)
    Next
    
'End If

On Error GoTo 0
Sheets("xxStdNames").Visible = False

End Sub
Sub StdzExtract(Optional UserExtract As Integer)


'On Error GoTo 0

    'save initial extract
    '================================================================================================================
    Set tmWB = ActiveWorkbook
    If Not UserExtract = 1 Then
    
        On Error Resume Next
        Rows("1:1").AutoFilter = False
        Cells.Columns.Hidden = False
        Range("A:A").Find(what:="x", lookat:=xlWhole).EntireRow.Delete Shift:=xlUp
        Range("A:A").Find(what:="No Data", lookat:=xlWhole).EntireRow.Delete Shift:=xlUp
        Range("A1:A" & FUN_lastrow("A")).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        
        Application.DisplayAlerts = False
        Set NewBook = Workbooks.Add
        tmWB.ActiveSheet.Copy Before:=NewBook.Sheets(1)
retrysave:
        If Dir(ZeusPATH & "\" & Usr & "_Initiative_Extract(" & FileName_PSC & ")" & ")" & svCNT & ".xlsx", vbDirectory) = vbNullString Then
            NewBook.SaveAs Filename:=ZeusPATH & "\" & Usr & "_Initiative_Extract(" & FileName_PSC & ")" & ")" & svCNT, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Else
            svCNT = svCNT + 1
            GoTo retrysave
        End If
        ActiveWorkbook.Close (False)
        tmWB.Activate
        Application.DisplayAlerts = True
        On Error GoTo 0
        
    Else
        On Error Resume Next
        Rows("1:1").AutoFilter = False
        Cells.Columns.Hidden = False
        On Error GoTo 0
    End If
    
'    'format SNA extract
'    '================================================================================================================
'    If UserExtract = 1 Then
'
'        'rearrange columns
'        '------------------------------------
'        'If UserExtract = 1 Then
'            Columns("AJ:AJ").Cut
'            Columns("Z:Z").Insert
'            Columns("AG:AG").Cut
'            Columns("AA:AA").Insert
'            Columns("AH:AH").Cut
'            Columns("AB:AB").Insert
'            Columns("AK:AK").Cut
'            Columns("AC:AC").Insert
'            Columns("AL:AL").Cut
'            Columns("AD:AD").Insert
'            Columns("AN:AN").Cut
'            Columns("AE:AE").Insert
'            Columns("AQ:AQ").Cut
'            Columns("AF:AF").Insert
'            Columns("AQ:AQ").Cut
'            Columns("AG:AG").Insert
'            Columns("AP:AP").Cut
'            Columns("AH:AH").Insert
'            Range("AI:AQ").Delete Shift:=xlToLeft
'        'End If
'        Columns("O:O").Copy
'        Columns("P:P").Insert
'        Columns("U:U").Delete Shift:=xlToLeft
'        Columns("W:X").Insert
'        Columns("AA:AA").Insert
'        Columns("AL:AL").Insert
'        Range("AL1").Value = "DATE_SUBMITTED"
'        Range("AL2:AL" & Range("A1").End(xlDown).Row).Value = ZeusForm.asscStartDate.Text & "-" & ZeusForm.asscEndDate.Text
'        Columns("Z:Z").ClearContents
'
'    End If
    
        
    'rearrange Columns
    '================================================================================================================
    Columns("Z:AH").Cut     'blue Nov section
    Columns("F:F").Insert
    Columns("M:M").Insert   'if UOM Qty field is not included
    Columns("P:P").Cut      'MID
    Columns("AN:AN").Insert
    Columns("Q:S").Delete   'SysNm, SysID, NetNm
    Columns("Y:Z").Insert   'prod sub-category, unique prod count
    Columns("AA:AA").Copy   'UOM qty 2
    Columns("AB:AB").Insert
    Columns("AE:AE").Cut 'EA price
    Columns("AH:AH").Insert
    Columns("AE:AE").Copy   'Reported Usage
    Columns("AE:AE").Insert
    Columns("AI:AI").Cut    'Reported spend
    Columns("AN:AN").Insert
    Columns("AI:AI").Delete 'Date submitted
    Columns("AI:AI").Cut   'Study Key
    Columns("AN:AN").Insert
    
    Debug.Assert (IsNumeric(Range("AA1").End(xlDown).Value))
    If IsNumeric(Range("AA1").End(xlDown).Value) = False Then MsgBox "You found it!  Please let Barry know.  Thank you."

    
    'Clean & Standardize
    '================================================================================================================
    
    'stdz mbrs
    '--------------------------------
    'Set mbrNm = Range("G2")
    'StdNmCol = "H"
'    If ZeusForm.NRsrch = True Then
'        On Error Resume Next
'        If Not Trim(Sheets("xxStdNames").Range("A3")) = "" Then
'            Set ID2NMrng = Range(Sheets("xxStdNames").Range("A3"), Sheets("xxStdNames").Range("A1").End(xlDown))
'            For Each mbr In Range(Range("AM2"), Range("AM2").End(xlDown))
'                Range("P" & mbr.Row).Value = ID2NMrng.Find(what:=mbr.Value, lookat:=xlWhole).Offset(0, 1).Value
'            Next
'        End If
'        Set NM2NMrng = Range(Sheets("xxStdNames").Range("C1"), Sheets("xxStdNames").Range("C1").End(xlDown))
'        For Each mbr In Range(Range("G2"), Range("G2").End(xlDown))
'            Range("H" & mbr.Row) = NM2NMrng.Find(what:=mbr.Value, lookat:=xlWhole).Offset(0, 1).Value
'        Next
'        On Error GoTo 0
'    Else
        Set M2Mrng = Range(Sheets("xxStdNames").Range("B1"), Sheets("xxStdNames").Range("B1").End(xlDown))
        Set M2Srng = Range(Sheets("xxStdNames").Range("D1"), Sheets("xxStdNames").Range("D1").End(xlDown))
        Set SYSrng = Range(Sheets("xxStdNames").Range("F1"), Sheets("xxStdNames").Range("F1").End(xlDown))
        If ZeusForm.NRSrch = True Then
            'Set mbrrng = Range(Range("P2"), Range("P1").End(xlDown))
            Range(Range("P2"), Range("P1").End(xlDown)).Select
        Else
            'Set mbrrng = Range(Range("AL2"), Range("AL1").End(xlDown))
            Range(Range("AL2"), Range("AL1").End(xlDown)).Select
        End If
        
        On Error Resume Next
        'For Each c In MbrRng
        For Each c In Selection
            Range("P" & c.Row).Value = M2Mrng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, -1).Value
            Range("P" & c.Row).Value = M2Srng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, -1).Value
            Range("P" & c.Row).Value = SYSrng.Find(what:=c.Offset(0, 3).Value, lookat:=xlWhole).Offset(0, -1).Value
        Next
        On Error GoTo 0
    'End If
    
    Range("AL:AM").Clear

    'Populate blank catnums
    '================================================================================================================
    'StdCatOffset = 23 (Col X)
    For Each c In Range(Range("A2"), Range("A1").End(xlDown)).Offset(0, 23)
        If Trim(c.Value) = "" Then
            If Not Trim(Range("W" & c.Row).Value) = "" Then
                c.Value = Trim(Range("W" & c.Row).Value)
            Else
                If Not Trim(Range("T" & c.Row).Value) = "" Then
                    c.Value = Trim(Range("T" & c.Row).Value)
                Else
                    If Not Trim(Range("Q" & c.Row).Value) = "" Then
                        c.Value = Trim(FUN_convCatnum(Range("Q" & c.Row).Value))
                    Else
                        uknCNT = uknCNT + 1
                        c.Value = "Unknown" & uknCNT
                    End If
                End If
            End If
        End If
    Next
    
    'Populate blank UOMs
    '================================================================================================================
    'UOMOffset = 26 (Col AA)
    For Each c In Range(Range("A2"), Range("A1").End(xlDown)).Offset(0, 26)
        If Trim(c.Value) = "" Or Trim(c.Value) = 0 Then
            c.Value = 0
            c.Offset(0, 1).Value = 1
        End If
    Next
    
    'stdz mftr names
    '================================================================================================================
    If TeamNm = "CSA" Then
        'StdMftrOffset = 21 (Col V)
        wbnm = ActiveWorkbook.Name
        Set stdmfgwb = Workbooks.Open("\\filecluster01\dfs\NovSecure2\SupplyNetworks\Analytics\DAT Resources\Supplier Name Standardization\Standard Manufacturer Names.xlsx")
        Workbooks(wbnm).Activate
        Set stdrng = Range(stdmfgwb.Sheets("Manufacturer Names").Range("A2"), stdmfgwb.Sheets("Manufacturer Names").Cells(stdmfgwb.Sheets("Manufacturer Names").Rows.Count, 1).End(xlUp))
        On Error Resume Next
        For Each c In Range(Range("A2"), Range("A2").End(xlDown)).Offset(0, 21)
            c.Value = stdrng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, 1).Value
            If c.Value = "" Then
                If c.Offset(0, -1).Value = "" Then
                    uknMftr = uknMftr + 1
                    c.Value = "Unknown" & uknMftr
                Else
                    c.Value = c.Offset(0, -1).Value
                End If
            End If
        Next
        stdmfgwb.Close (False)
    End If


Exit Sub
'::::::::::::::::::::::::::::::::::::::




End Sub
Sub Import_Spend()

'from create report
'from indiv
    'from extract
    'from interface
        'from VUN
        'from RDM

    'Rollup !Must be sorted by catnmbr, and txt to col and general format!
    '====================================================================================================
    Application.StatusBar = "Sorting for ASF import..."
    LastRow = FUN_lastrow(1)
    Range(Range("A2"), Range("AK" & LastRow)).Interior.ColorIndex = 0
    Call FUN_Sort(ActiveSheet.Name, Range("A2:AK" & LastRow), Range("X2:X" & LastRow), 1, Range("AD2:AD" & LastRow), 1)

    'capture line item IDs
    '====================================================================================================
    Call FUN_TestForSheet("xxProgramData")
    Sheets("xxProgramData").Range("A1").Value = "Original Sort IDs (not rolled up)"
    Sheets("xxProgramData").Range("A2").Value = "x"
    Range(Sheets("Spend search").Range("A2"), Sheets("Spend search").Range("A" & LastRow)).Copy
'    If Trim(Sheets("xxProgramData").Range("A2").Value) = "" Then
'        Sheets("xxProgramData").Range("A2").PasteSpecial Paste:=xlPasteValues
'    Else
        Sheets("xxProgramData").Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
'    End If
    Sheets("xxProgramData").Visible = False
    Sheets("Spend Search").Select

    'format
    '====================================================================================================
    Application.DisplayAlerts = False
    On Error Resume Next
    'Catnum
    'Range("X2:X" & lastrow + 2).NumberFormat = "@"     '<--Format as text
'    Range("X2:X" & lastrow + 2).TextToColumns Destination:=Range("X2"), DataType:=xlDelimited, _
'        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'        :=Array(1, 1), TrailingMinusNumbers:=True
    'UOM qty
    Range("AA2:AA" & LastRow + 2).TextToColumns Destination:=Range("AA2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'UOM qty
    Range("AB2:AB" & LastRow + 2).TextToColumns Destination:=Range("AB2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'price per UOM
    Range("AD2:AD" & LastRow + 2).TextToColumns Destination:=Range("AD2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'reported usage
    Range("AE2:AE" & LastRow + 2).TextToColumns Destination:=Range("AE2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'annualized usage
    Range("AF2:AF" & LastRow + 2).TextToColumns Destination:=Range("AF2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'annualized ea usage
    Range("AG2:AG" & LastRow + 2).TextToColumns Destination:=Range("AG2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'EA Price
    Range("AH2:AH" & LastRow + 2).TextToColumns Destination:=Range("AH2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'Total Spend
    Range("AK2:AK" & LastRow + 2).TextToColumns Destination:=Range("AK2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'PIM
    'Range("E2:E" & lastrow + 2).NumberFormat = "@"
'    Range("E2:E" & lastrow + 2).TextToColumns Destination:=Range("E2"), DataType:=xlDelimited, _
'        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'        :=Array(1, 1), TrailingMinusNumbers:=True
    'Range("E2:E" & lastrow + 2 & ", N2:P" & lastrow + 2 & ", S2:T" & lastrow + 2).NumberFormat = "General"
    On Error GoTo 0
    
    'Roll up
    '====================================================================================================
    Range("X2").Select
    Do
        mtchcats = Application.CountIf(Range("X:X"), ActiveCell.Value)
        Set AllCats = Range(ActiveCell, ActiveCell.Offset(mtchcats - 1, 0))
        
        For Each itm In AllCats
        
            If Not Range("AL" & itm.Row).Value = 1 Then
                For Each c In AllCats
                    If c.Address = itm.Address Or Range("AL" & c.Row).Value = 1 Then GoTo nxtC          '(if not same item and item not already added to another)
                    If Not Range("AD" & c.Row).Value = Range("AD" & itm.Row).Value Then GoTo nxtC         '(if same Unit price)
                    If Not Range("AC" & c.Row).Value = Range("AC" & itm.Row).Value Then GoTo nxtC         '(if same Pkg Desc)
                    If Not Range("AA" & c.Row).Value = Range("AA" & itm.Row).Value Then GoTo nxtC         '(if same UOM)
                    If Not Range("V" & c.Row).Value = Range("V" & itm.Row).Value Then GoTo nxtC         '(if same mfg)
                    If Not Range("P" & c.Row).Value = Range("P" & itm.Row).Value Then GoTo nxtC         '(if same member)
                    Range("AE" & itm.Row).Value = Range("AE" & itm.Row).Value + Range("AE" & c.Row).Value  '(add Reported usage)
                    Range("AF" & itm.Row).Value = Range("AF" & itm.Row).Value + Range("AF" & c.Row).Value  '(add annualized usage)
                    Range("AG" & itm.Row).Value = Range("AG" & itm.Row).Value + Range("AG" & c.Row).Value  '(add annualized EA usage)
                    Range("AK" & itm.Row).Value = Range("AK" & itm.Row).Value + Range("AA" & c.Row).Value * Range("AF" & c.Row).Value * Range("AH" & c.Row).Value  '(spend + calculated spend(UOM*Annual Usage*EA Price))
                    Range("AL" & c.Row).Value = 1
                    DoEvents
                    Application.StatusBar = "Rolling up similar items: " & c.Row
nxtC:           Next
            End If
        Next
 
        ActiveCell.Offset(mtchcats, 0).Select
    Loop Until Trim(ActiveCell.Value) = ""
    
    Range("P:P, V:V, X:X, AA:AA, AC:AD").EntireColumn.Interior.Color = 65535
    Range("AE:AG,AK:AK").EntireColumn.Interior.Color = 65280
    
    On Error GoTo errhndlNoRollups
    Range("AL:AL").SpecialCells(xlCellTypeConstants, 1).EntireRow.Select
    On Error GoTo 0
    Selection.Interior.ColorIndex = 3
    
    'Save original extract
    '====================================================================================================
    Application.DisplayAlerts = False
    wbnm = ActiveWorkbook.Name
    Set NewBook = Workbooks.Add
    Workbooks(wbnm).Sheets("Spend Search").Copy Before:=NewBook.Sheets(1)
    Application.StatusBar = "Saving ASF Rollups..."
retrysave:
    If Dir(ZeusPATH & "\ASFrollups(" & FileName_PSC & ")" & ")" & svCNT & ".xlsx", vbDirectory) = vbNullString Then
        NewBook.SaveAs Filename:=ZeusPATH & "\ASFrollups(" & FileName_PSC & ")" & ")" & svCNT, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Else
        svCNT = svCNT + 1
        GoTo retrysave
    End If
    ActiveWorkbook.Close (False)
    
    If Not Application.CountIf(Range("AL:AL"), 1) = 0 Then
        Call FUN_Sort(ActiveSheet.Name, Range("A2:AL" & LastRow), Range("AL2:AL" & LastRow), 2)
        Range("AL2:AL" & Range("AL" & LastRow).End(xlUp).Row).EntireRow.Select
        Selection.Delete Shift:=xlUp
    End If
    
noRollups:
    'clean extract
    '====================================================================================================
    Call METH_ClinicalClean     '>>>>>>>>>>
    If ZeusForm.extractSrch = True Then Call StandardizeMfg       '>>>>>>>>>>

    'Import to TM
    '===================================================================================================
    '===================================================================================================
    Sheets("Spend Search").Select
    LastRow = FUN_lastrow(1)
    Range(Range("A2"), Range("AK" & LastRow)).Font.Name = "Arial"
    Range(Range("A2"), Range("AK" & LastRow)).Font.Size = 8
    Range(Range("A2"), Range("AK" & LastRow)).Borders.LineStyle = xlNone
    Call FUN_TestForSheet("Line Item Data")
    
    If SetupSwitch = 2 Then

        If Not Trim(Range("A6").Value) = "" Then Range(Range("A6"), Range("A" & LastRow + 4)).EntireRow.Insert Shift:=xlDown
        
        Range(Sheets("Spend search").Range("A2"), Sheets("Spend search").Range("AK" & LastRow)).Copy
        Range("A6").PasteSpecial xlPasteAll
        
        'mbr data
        '---------------------
        Range("Z5").AutoFill Destination:=Range("Z5:Z" & LastRow + 4)   'Unique Prod count
        Range("AG5:AH5").AutoFill Destination:=Range("AG5:AH" & LastRow + 4) 'EA price
        Range("AJ5").AutoFill Destination:=Range("AJ5:AJ" & LastRow + 4) 'calculated spend
        Range("AL5").AutoFill Destination:=Range("AL5:AL" & LastRow + 4) 'calculated spend
        
        'plvling
        '---------------------
        Range("AN5:AQ5").AutoFill Destination:=Range("AN5:AQ" & LastRow + 4)

        'Bench
        '---------------------
        Range("AR5:BF5").AutoFill Destination:=Range("AR5:BF" & LastRow + 4)
        'Range("AR6:BF" & lastrow + 4).ClearFormats
'        Range("AR6:AU" & lastrow + 4).NumberFormat = "$#,##0.00"
'        Range("AW6:AZ" & lastrow + 4).NumberFormat = "$#,##0.00"
'        Range("BB6:BE" & lastrow + 4).NumberFormat = "$#,##0.00"
'        Range("AV6:AV" & lastrow + 4).NumberFormat = "%#,##0"
'        Range("BA6:BA" & lastrow + 4).NumberFormat = "%#,##0"
'        Range("BF6:BF" & lastrow + 4).NumberFormat = "%#,##0"
        
        'Supplier
        '---------------------
        Range(Range("BG5"), Range("BF5").Offset(0, suppNMBR * 30)).AutoFill Destination:=Range(Range("BG5"), Range("BF5").Offset(LastRow - 1, suppNMBR * 30))
        'Range(Range("BG6"), Range("BF6").Offset(lastrow - 1, suppNMBR * 30)).ClearFormats
'        For i = 1 To supplier
'            Range("AR6:AU" & lastrow + 4).NumberFormat = "$#,##0.00"
'            Range("AW6:AZ" & lastrow + 4).NumberFormat = "$#,##0.00"
'            Range("BB6:BE" & lastrow + 4).NumberFormat = "$#,##0.00"
'            Range("AV6:AV" & lastrow + 4).NumberFormat = "%#,##0"
'            Range("BA6:BA" & lastrow + 4).NumberFormat = "%#,##0"
'            Range("BF6:BF" & lastrow + 4).NumberFormat = "%#,##0"
'        Next

        If Not CreateReport = True Then
            Range(Range("AI6"), Range("AI5").Offset(LastRow - 1, 0)).Interior.Color = 65535
            Range("A6:A" & LastRow + 4).EntireRow.Calculate
        End If
        Range(Range("A6"), Range("BG6").Offset(LastRow + 4, (suppNMBR * 30) - 1)).Font.Size = 8
        If Trim(Range("A5").Value) = "" Then Range("A5").EntireRow.Delete Shift:=xlUp 'Range(Range("A5"), Range("BG5").Offset(0, (suppNMBR * 30) - 1)).Delete Shift:=xlUp
        'ItmNmbr = Application.CountA(Sheets("Line Item Data").Range("A5:A" & Range("A4").End(xlDown).Row)) - 3
        ItmNmbr = Application.CountA(Range("X:X")) - 1
        'ItmNmbr = 3
        
        'Create owners Column for CAHN and MNS
        '===================================================================================================
        If OwnrNmbr > 0 Then
            
            Call label_owner_items
            
'            For supp = 1 To 10
'                For Each c In Range(ConvBKMRK.Offset((MbrNMBR + 1) + (MbrNMBR + 8) * (supp - 1), 1), ConvBKMRK.Offset((MbrNMBR + 1) + (MbrNMBR + 8) * (supp - 1), 0).End(xlToRight))
'                    c.Formula = Replace(c.Formula, c.Offset(-1, 0).Address(0, 0), c.Offset(-OwnrNmbr - 1, 0).Address(0, 0))
'                Next
'            Next
            
'            'Fix unique product count for owners
'            '-----------------------
'            OwnrCol.Offset(1, 1).Formula = Sheets("Line Item Data").Range("Z5").Formula
'            OwnrCol.Offset(1, 1).Replace what:="$P", Replacement:="$MU", lookat:=xlPart
'            OwnrCol.Offset(1, 1).AutoFill Destination:=Range(OwnrCol.Offset(1, 1), OwnrCol.Offset(lastrow - 1, 1))
'
            On Error GoTo 0
        End If
        
        'Misc Ending Functions
        '===================================================================================================
        If NtwkSource = "Extract" Then Call StdzContractedMftrs        '>>>>>>>>>>
        Call Calculate_Priceleveling    '>>>>>>>>>>
        
    Else
        If Trim(Range("A1").Value) = "" Then
            Range(Sheets("Spend search").Range("A1"), Sheets("Spend search").Range("AK" & LastRow)).Copy
            Range("A1").PasteSpecial xlPasteAll
            Rows("1:1").Interior.Color = 10921638
            Rows("2:2").Select
            ActiveWindow.FreezePanes = True
        Else
            Range(Sheets("Spend search").Range("A2"), Sheets("Spend search").Range("AK" & LastRow)).Copy
            Range("A1").End(xlDown).Offset(1, 0).PasteSpecial xlPasteAll
        End If
        
        Range("A" & LastRow).Select
    End If
    
    Range("V2:AC" & LastRow + 4, "AE2:AG" & LastRow + 4).HorizontalAlignment = xlCenter
    Range("AI2:AI" & LastRow + 4).HorizontalAlignment = xlCenter
    Range("AD:AD, AH:AH, AK:AK").NumberFormat = "$#,##0.00"
   
    
    Application.StatusBar = False


Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoRollups:
On Error GoTo 0
Resume noRollups


End Sub
   
