Attribute VB_Name = "B__Import"
'Updated Arrays
'-----------------
'Public NetworkNamesParent() As Variant
Dim CreateXrefWB As Workbook
'Public NetworkArrayParent() As Variant

'Extract Arrays
'-----------------
'Public IDtoNM() As String
'Public NMtoNM() As String
'Public DateRange As String
'Public NetworkRun As String
'Public DateRanges() As String
'Public NetworkRuns() As String
'Public PRSNameParent() As String
'Public PRSArrayParent() As Variant
'Public PRSinit As Integer


Sub Working_File_Main()

    
    If Not FUN_Save = vbYes Then Exit Sub
    Set wfWB = Workbooks.Add
    SetupSwitch = FUN_SetupSwitch '(new workbook makes this = 3)
    
    Call ContractInfo_Template
    On Error Resume Next
    Sheets("Sheet1").Visible = False
    Sheets("Sheet2").Visible = False
    Sheets("Sheet3").Visible = False
    On Error GoTo 0
    MainCall = 1
    
    Sheets("contract info").Range("N1").Value = ZeusForm.asscNetwork.Value
    Sheets("contract info").Range("N2").Value = ZeusForm.asscPSC.Value
    'Call Import_TierInfo
    'Call Import_StdznIndex
    Call Import_Pricefile
    Call Import_UNSPSC
    Call Import_PRS
    'Call Import_Scopeguide
    Call Import_CheatSheet
    Sheets("pricefile keywords").Select 'must be the active sheet or else pricefile descs will be erased
    Call KeywordGenerator

    'save As Workbook
    '============================================================================================
    wfWB.Activate
    Application.DisplayAlerts = False
    wfWB.SaveAs Filename:=ZeusPATH & "Working File(" & FileName_PSC & ")", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    
    Sheets("Contract info").Select
    Range("A1").Select
    MainCall = 0
    Call FUN_CalcBackOn
    
    
End Sub
Sub Import_TierInfo()
    
    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset
    Dim QCTierArray() As String

    If Not QCFlg = True Then
        If SetupSwitch = 2 Then
            Sheets("index").Visible = True
            Sheets("index").Select
            'Range(ConTblBKMRK.Offset(0, 1), ConTblBKMRK.End(xlDown).Offset(0, 10)).ClearContents
            Range(ConTblBKMRK.Offset(9, 0), ConTblBKMRK.End(xlDown).Offset(0, 10)).FormatConditions.Delete
            Range(ConTblBKMRK.Offset(9, 0), ConTblBKMRK.End(xlDown).Offset(0, 10)).Delete Shift:=xlUp
            Range(ConTblBKMRK.Offset(0, 1), ConTblBKMRK.End(xlDown).Offset(0, 10)).Delete Shift:=xlToLeft
        Else
            Call FUN_TestForSheet("Contract Info")
            If Not (Range("M2").Value = "Network:" And Range("M2").Value = "PSC:") Then Call ContractInfo_Template
            Range(Sheets("contract info").Range("B1"), Sheets("contract info").Range("K1").Offset(59, 0)).ClearContents
        End If
    End If
    
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    For i = 1 To suppNMBR
        If QCFlg = True Then ReDim QCTierArray(1 To 1)
        
        'Build Query string
        '============================================================================================================
        connmbr = ZeusForm.asscContracts.List(i - 1)
        
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        QueryPATH = FUN_ConvTags(AdminconfigStr, "TierInfoQuery_Path")
        Set fileObj = objFSO.GetFile(QueryPATH)
        sqlstr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
        sqlstr = Replace(sqlstr, "!!COMPANYCODE!!", CmpyCD)
        sqlstr = Replace(sqlstr, "!!CONTRACTID!!", connmbr)
        
'        SELECTstr = "SELECT DISTINCT OV.VENDOR_NAME, OC.CONTRACT_NUMBER, (cpd.First_name + ' ' + cpd.last_name) AS PORTFOLIO_EXECUTIVE, OC.CONTRACT_EXP_DATE, OC.CONTRACT_EFF_DATE, cav.Attribute_value_name, OC.AGREEMENT_TYPE_KEY AS NOVAPLUS, OPT.TIER_DESCRIPTION, OPT.LABEL"
'        FROMstr = " FROM OCSDW_CONTRACT OC INNER JOIN OCSDW_VENDOR AS OV ON OC.VENDOR_KEY = OV.VENDOR_KEY  INNER JOIN OCSDW_PRICE_Novation AS OP ON OC.CONTRACT_ID = OP.CONTRACT_ID AND OP.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "' INNER JOIN OCSDW_PRICE_TIER AS OPT ON OP.PRICE_TIER_KEY = OPT.PRICE_TIER_KEY  INNER JOIN OCSDW_CONTRACT_PROGRAM_DETAIL cpd ON OC.CONTRACT_ID = cpd.CONTRACT_ID LEFT OUTER JOIN OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL cav ON cav.CONTRACT_ID = oc.CONTRACT_ID AND cav.attribute_value_id = '963'"
'        WHEREstr = " WHERE OC.CONTRACT_NUMBER ='" & connmbr & "' ORDER BY OPT.LABEL" 'AND (OC.Status_Key = 'ACTIVE' or OC.Status_Key = 'signed' or OC.Status_Key = 'expired')
'        SQLstr = SELECTstr & FROMstr & WHEREstr
    
        'execute query & import
        '============================================================================================================
        On Error GoTo errhndlNORECSET
        recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
        recset.MoveFirst
        On Error GoTo 0
        
        'import contract info
        '-------------------------
        If Not QCFlg = True Then
            If SetupSwitch = 2 Then
                'If not instr(Sheets("Line item Data").Range("BG4").Offset(0, (i - 1) * 30).Value,"Proposed Catalog #") >0 Then Call Add_Scenario(connmbr)
                Sheets("Line item Data").Range("BG3").Offset(0, (i - 1) * 30).Value = recset.Fields(0).Value
                ConTblBKMRK.Offset(0, i).Value = recset.Fields(0).Value    'name
                ConTblBKMRK.Offset(1, i).Value = connmbr                   'number
                ConTblBKMRK.Offset(2, i).Value = recset.Fields(2).Value    'portfolio exec
                ConTblBKMRK.Offset(3, i).Value = recset.Fields(4).Value    'con start
                ConTblBKMRK.Offset(4, i).Value = recset.Fields(3).Value    'con end
                If recset.Fields(5).Value = "STANDARDIZATION" Then
                    ConTblBKMRK.Offset(5, i).Value = "ü"
                    ConTblBKMRK.Offset(5, i).Font.Color = 11250945
                Else
                    ConTblBKMRK.Offset(5, i).Value = "û"
                    ConTblBKMRK.Offset(5, i).Font.Color = 20223
                End If
                If recset.Fields(6).Value = "NOVAPLUS" Then
                    ConTblBKMRK.Offset(6, i).Value = "ü"
                    ConTblBKMRK.Offset(6, i).Font.Color = 11250945
                Else
                    ConTblBKMRK.Offset(6, i).Value = "û"
                    ConTblBKMRK.Offset(6, i).Font.Color = 20223
                End If
            Else
                Sheets("contract info").Range("A1").Offset(0, i).Value = recset.Fields(0).Value
                Sheets("contract info").Range("A2").Offset(0, i).Value = connmbr
                Sheets("contract info").Range("A3").Offset(0, i).Value = recset.Fields(2).Value
                Sheets("contract info").Range("A4").Offset(0, i).Value = recset.Fields(4).Value
                Sheets("contract info").Range("A5").Offset(0, i).Value = recset.Fields(3).Value
                If recset.Fields(5).Value = "STANDARDIZATION" Then
                    Sheets("contract info").Range("A6").Offset(0, i).Value = "Yes"
                Else
                    Sheets("contract info").Range("A6").Offset(0, i).Value = "No"
                End If
                If recset.Fields(6).Value = "NOVAPLUS" Then
                    Sheets("contract info").Range("A7").Offset(0, i).Value = "Yes"
                Else
                    Sheets("contract info").Range("A7").Offset(0, i).Value = "No"
                End If
            End If
        End If
        
        'impor tiers
        '-------------------------
        For j = 1 To recset.RecordCount
            tiernum = recset.Fields(8).Value
            If Not QCFlg = True Then
                If SetupSwitch = 2 Then
                    If tiernum > maxtier Then
                        Range(ConTblBKMRK.Offset(9 + maxtier, 0), ConTblBKMRK.Offset(8 + tiernum, 0)).EntireRow.Insert
                        maxtier = tiernum
                    End If
                    'Set CurrTier = ConTblBKMRK.Offset(Replace(tiernum, "00", "") + 8, i)
                    Set CurrTier = ConTblBKMRK.Offset(tiernum + 8, i)
                Else
                    'Set CurrTier = Sheets("contract info").Range("A1").Offset(Replace(tiernum, "00", "") + 6, i)    'must find corresponding tier incase tiers are missing
                    Set CurrTier = Sheets("contract info").Range("A1").Offset(tiernum + 6, i)    'must find corresponding tier incase tiers are missing
                End If
                
                If Trim(CurrTier) = "" Then
                    CurrTier.Value = recset.Fields(7).Value
                Else
                    CurrTier.Value = CurrTier.Value & "; " & recset.Fields(7).Value  'if multiple price programs then must combine with what's already there
                End If
            Else
                ReDim Preserve QCTierArray(1 To j)
                If Trim(QCTierArray(tiernum)) = "" Then
                    QCTierArray(tiernum) = recset.Fields(7).Value
                Else
                    QCTierArray(tiernum) = QCTierArray(tiernum) & "; " & recset.Fields(7).Value
                End If
            End If
            
            recset.MoveNext
        Next
        
        'if QCcheck then compare array to current tier cells
        '-------------------------
        If QCFlg = True Then
            For QCtier = 1 To UBound(QCTierArray)
                Set ReportTier = ConTblBKMRK.Offset(QCtier + 8, i)
                If Not QCTierArray(QCtier) = ReportTier.Value Then
                    QCChkFlg = False
                    ReportTier.Interior.Color = 16711935
                ElseIf ReportTier.Interior.Color = 16711935 Then
                    ReportTier.Interior.Color = 0
                End If
            Next
        End If
        
        'items on contract formula
        '-------------------------
        If Not QCFlg = True Then
            
            If SetupSwitch = 2 Then
                pfcount = 0
                For Each sht In ActiveWorkbook.Sheets
                    If InStr(sht.Name, "Pricing") Then
                        pfcount = pfcount + 1
                        If pfcount = i Then
                            ConTblBKMRK.Offset(7, i).Formula = "=counta('" & sht.Name & "'!A:A)-1"
                            Exit For
                        End If
                    End If
                Next
                
                ConTblBKMRK.Offset(8, i).Value = "Best Price"
                
                'reattach supplier names in formulas
                '------------------------------------
                suppnmAddress = "Index!" & ConTblBKMRK.Offset(0, i).Address
                pftab = "'" & FUN_SuppName(i) & " Pricing'!"
                MSGraphBKMRK.Offset(0, i * 2 - 1).Formula = "=" & suppnmAddress
                prsBKMRK.Offset(0, i * 2 - 1).Formula = "=" & suppnmAddress
                prsBKMRK.Offset(0, i * 2).Formula = "=concatenate(" & suppnmAddress & ", "" Reported"")"
                'NonConBKMRK.Offset((MbrNMBR + 8) * (i - 1) - 2, 0).Formula = "=CONCATENATE(" & suppnmAddress & ","" "",Index!" & ConTblBKMRK.Offset(1, i).Address & ","" - "",IF(Index!" & ConTblBKMRK.Offset(8, i).Address & "=""Best Price"",OFFSET(" & pftab & "$J$1,0,MATCH(MIN(" & pftab & "$K$2:$U$2)," & pftab & "$K$2:$U$2,0)),Index!" & ConTblBKMRK.Offset(8, i).Address & "))"
                'ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1) - 2, 0).Formula = "=CONCATENATE(" & suppnmAddress & ","" "",Index!" & ConTblBKMRK.Offset(1, i).Address & ","" - "",IF(Index!" & ConTblBKMRK.Offset(8, i).Address & "=""Best Price"",OFFSET(" & pftab & "$J$1,0,MATCH(MIN(" & pftab & "$K$2:$U$2)," & pftab & "$K$2:$U$2,0)),Index!" & ConTblBKMRK.Offset(8, i).Address & "))"
                NonConBKMRK.Offset((MbrNMBR + 8) * (i - 1), 6).Formula = "=CONCATENATE(""Savings % of Spend With ""," & suppnmAddress & ")"
                ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1), 9).Formula = "=CONCATENATE(""# of Unique Products Not matched to ""," & suppnmAddress & ","" Contract"")"
    
                'HCO supplier sections
                '--------------------------
                'Sheets("Line Item Data").Range("BG1").Offset(0, (i - 1) * 31).Value = Range(Bkmrk).Offset(0, i).Value      'name
                'Sheets("Line Item Data").Range("BG1").Offset(0, (i - 1) * 31 + 1).Value = ConTblBKMRK.Offset(1, i).Value   'contract
                'Sheets("Line Item Data").Range("BG1").Offset(0, (i - 1) * 31 + 2).Value = "Tier " & Application.CountA(Range(ConTblBKMRK.Offset(9, i), ConTblBKMRK.Offset(9, i).End(xlDown)))
            End If
            
        End If
        
        recset.Close
nxtCon:
    Next

    'formatting
    '------------------------
    If Not QCFlg = True Then
        
        If SetupSwitch = 2 Then
            Set DataRng = Range(ConTblBKMRK.Offset(0, 1), ConTblBKMRK.Offset(maxtier + 8, suppNMBR))
            
            DataRng.Font.Size = 8
            Range(ConTblBKMRK.Offset(5, 1), ConTblBKMRK.Offset(6, suppNMBR)).Font.Size = 22
            Range(ConTblBKMRK.Offset(5, 1), ConTblBKMRK.Offset(6, suppNMBR)).Font.Name = "Wingdings"
            DataRng.HorizontalAlignment = xlCenter
            DataRng.VerticalAlignment = xlTop
            DataRng.WrapText = True
            DataRng.EntireRow.AutoFit
            'Range(ConTblBKMRK.Offset(maxtier + 9, 0), ConTblBKMRK.End(xlDown).Offset(0, 10)).Clear
    
            'Range(ConTblBKMRK.Offset(9, 0), ConTblBKMRK.End(xlDown)).HorizontalAlignment = xlRight
            DataRng.Borders.LineStyle = xlContinuous
            DataRng.Borders.ColorIndex = 1
            'Range(ConTblBKMRK.Offset(1, 1), ConTblBKMRK.Offset(maxtier + 8, suppNMBR)).Interior.color = 65535
            'Range(ConTblBKMRK.Offset(0, 1), ConTblBKMRK.Offset(0, suppNMBR)).Interior.Color = 10040115
            For i = 1 To maxtier
                ConTblBKMRK.Offset(8 + i, 0).Value = "Tier " & i & " Requirements"
            Next
            Set labelsRng = Range(ConTblBKMRK, ConTblBKMRK.End(xlDown))
            labelsRng.Borders.LineStyle = xlContinuous
            labelsRng.Borders.Color = 12566463
        Else
            Sheets("contract info").Range("B2:K7").HorizontalAlignment = xlCenter
        End If
        
    End If
    
    On Error Resume Next
    Set conn = Nothing
    Set recset = Nothing
    On Error GoTo 0
    
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNORECSET:
If Not QCFlg = True Then

    If SetupSwitch = 2 Then
        ConTblBKMRK.Offset(9, i).Value = "No Tier Info found"
        ConTblBKMRK.Offset(1, i).Value = connmbr
    Else
        Sheets("contract info").Range("A8").Offset(0, i).Value = "No Tier Info found"
        Sheets("contract info").Range("A2").Offset(0, i).Value = connmbr
    End If
    
End If
Set recset = Nothing
Resume nxtCon


End Sub
Sub Import_Pricefile(Optional pfsingle As Integer, Optional RefreshFlg As Boolean)

'[TBD] some tiers have 0 and some have 0 for all tiers
'[TBD] bring in all UOM levels out to the side to use with Einstein

    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset
    Dim conpos As Integer
    Dim i As Integer
    Dim parseName As String
    
    'Setup
    '===================================================================================================================================================================
    If pfsingle = 1 Then
        pfcount = 1
    Else
        pfcount = suppNMBR
        If Not SetupSwitch = 2 Then
            Call FUN_TestForSheet("Pricefile Keywords")
            Cells.Clear
            Range("A1:A2").Value = "x"
        End If
        
        'Delete blank pricing tabs
        '-----------------------------
        For Each sht In ActiveWorkbook.Sheets
            'If Left(sht.Name, 8) = "Supplier" And Right(sht.Name, 8) = " Pricing" Then sht.Delete
'            If Right(sht.Name, 8) = " Pricing" Then
'                sht.Delete
            If Right(sht.Name, 16) = " Cross Reference" Then
                If Application.CountA(sht.Range("A:A")) - 1 = 0 Then Sheets(Replace(sht.Name, "Cross Reference", "Pricing")).Cells.Clear
            End If
        Next
  
    End If
    
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    For i = 1 To pfcount
        
        'Rename Pricing and xref tabs
        '===================================================================================================================================================================
        On Error GoTo 0
        If pfsingle = 1 Then
            connmbr = Trim(Application.InputBox(prompt:="Contract number:", Type:=2))
'            For con = 1 To suppNMBR
'                If connmbr = ZeusForm.asscContracts.List(con - 1) Then OnCon = True
'            Next
        Else
            connmbr = ZeusForm.asscContracts.List(i - 1)
            'OnCon = True
        End If
        
        If SetupSwitch = 2 Then
            
            On Error GoTo ERR_StandAlone
            conpos = ConTblBKMRK.Offset(1, 0).EntireRow.Find(what:=connmbr).Column - ConTblBKMRK.Column
            On Error GoTo 0
            
            PricingNm = Sheets("Line item data").Range("BG3").Offset(0, (conpos - 1) * 30).Value
            If Trim(PricingNm) = "" Then
                Call Add_Scenario(connmbr)
            Else
                SuppNm = FUN_SuppName(conpos)
                If SuppNm = "#REF!" Then
                    Call Add_Scenario(connmbr)
                Else
                    Call FUN_TestForSheet(SuppNm & " Cross Reference")
                    If ActiveSheet.Name = "Supplier" & conpos & " Cross Reference" Or RefreshFlg = True Then
                        On Error GoTo ERR_LongName
                        ActiveSheet.Name = PricingNm & " Cross Reference"
                        longerr = 0
                        On Error GoTo 0
                    End If
                    Call FUN_TestForSheet(SuppNm & " Pricing")
                    If ActiveSheet.Name = "Supplier" & conpos & " Pricing" Or RefreshFlg = True Then ActiveSheet.Name = PricingNm & " Pricing"
                    ActiveSheet.AutoFilterMode = False
                    Cells.Clear
                End If
            End If
        Else
            On Error GoTo ERR_StandAlone
            Sheets(connmbr & " Pricing").Visible = True
            Sheets(connmbr & " Pricing").Select
            On Error GoTo 0
            Cells.Clear
        End If
        
NotPres:
        'Build Query string
        '===================================================================================================================================================================
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        QueryPATH = FUN_ConvTags(AdminconfigStr, "PricefileQuery_Path")
        Set fileObj = objFSO.GetFile(QueryPATH)
        sqlstr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
        sqlstr = Replace(sqlstr, "!!CONTRACTID!!", connmbr)
        sqlstr = Replace(sqlstr, "!!COMPANYCODE!!", CmpyCD)
  
'        SELECTstr = "SELECT DISTINCT OPD.VENDOR_PRODUCT_NUMBER AS Product_ID, OP.UOM_KEY AS Contract_UOM, OPD.PRODUCT_DESCRIPTION AS Product_Name, 'DEFAULT' AS Base_UOM, ISNULL(OPK.LEVEL_1_UOM_Qty,1) * ISNULL(OPK.LEVEL_2_UOM_Qty,1) * ISNULL(OPK.LEVEL_3_UOM_Qty,1) AS Base_UOM_Quantity, prog.Program_Name, 'DEFAULT' AS Price_Type, CONVERT(nvarchar(30), OP.PRICE_EFF_DATE, 101) AS Price_Start, CONVERT(nvarchar(30), OP.PRICE_EXP_DATE, 101) AS Price_End, CAST(OP.PRICE as MONEY) AS Tier_Price, OPT.LABEL AS Tier, OP.PRICE_STATUS_CODE"
'        FROMstr = " FROM OCSDW_CONTRACT AS OC INNER JOIN OCSDW_VENDOR AS OV ON OC.VENDOR_KEY = OV.VENDOR_KEY INNER JOIN OCSDW_PRICE_Novation AS OP ON OC.CONTRACT_ID = OP.CONTRACT_ID INNER JOIN OCSDW_PRODUCT AS OPD ON OPD.PRODUCT_KEY = OP.PRODUCT_KEY INNER JOIN OCSDW_PRICE_TIER AS OPT ON OP.PRICE_TIER_KEY = OPT.PRICE_TIER_KEY INNER JOIN OCSDW_PROGRAM AS PROG ON PROG.PROGRAM_KEY = OPT.PROGRAM_KEY"
'        SubSELECTstr = " INNER JOIN (SELECT DISTINCT LEVEL_1.PRODUCT_KEY, LEVEL_1_UOM_KEY, LEVEL_1.LEVEL_1_UOM_Qty, LEVEL_2_UOM_KEY, LEVEL_2.LEVEL_2_UOM_Qty, LEVEL_3_UOM_KEY, LEVEL_3.LEVEL_3_UOM_Qty"
'        SubFROMstr = " FROM (SELECT PRODUCT_KEY, FROM_UOM_KEY AS LEVEL_1_UOM_KEY, CAST(CONVERSION_RATE AS real) AS LEVEL_1_UOM_Qty FROM OCSDW_PRODUCT_PACKAGE AS PKG WHERE (PACKAGE_LEVEL LIKE '1%')) AS LEVEL_1"
'        SubJOIN1str = " FULL OUTER JOIN (SELECT PRODUCT_KEY, FROM_UOM_KEY AS LEVEL_2_UOM_KEY, CAST(CONVERSION_RATE AS real) AS LEVEL_2_UOM_Qty FROM OCSDW_PRODUCT_PACKAGE AS PKG WHERE (PACKAGE_LEVEL LIKE '2%')) AS LEVEL_2 ON LEVEL_1.PRODUCT_KEY = LEVEL_2.PRODUCT_KEY"
'        SubJOIN2str = " FULL OUTER JOIN (SELECT PRODUCT_KEY, FROM_UOM_KEY AS LEVEL_3_UOM_KEY, CAST(CONVERSION_RATE AS real) AS LEVEL_3_UOM_Qty FROM OCSDW_PRODUCT_PACKAGE AS PKG WHERE (PACKAGE_LEVEL LIKE '3%')) AS LEVEL_3 ON LEVEL_1.PRODUCT_KEY = LEVEL_3.PRODUCT_KEY) AS OPK ON OPD.PRODUCT_KEY = OPK.PRODUCT_KEY"
'        WHEREstr = " WHERE (OC.CONTRACT_NUMBER='" & connmbr & "') AND (OP.PRICE_STATUS_CODE = 'ACTIVE' OR OP.PRICE_STATUS_CODE = 'FUTURE') AND OP.COMPANY_CODE = '" & ZeusForm.AsscCompany.Value & "' ORDER BY Product_ID, Program_Name, Price_Status_code, Tier;"
'
'        SQLstr = SELECTstr & FROMstr & SubSELECTstr & SubFROMstr & SubJOIN1str & SubJOIN2str & WHEREstr

        'execute query & import
        '===================================================================================================================================================================
        Application.StatusBar = "Downloading Pricefile..."
        On Error GoTo errhndlNORECSET
        recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
        recset.MoveFirst
        On Error GoTo 0
        
        Range("A2").CopyFromRecordset recset
        
        'setup tier headers
        '-----------------------
        For j = 0 To 8
            Range("A1").Offset(0, j).Value = recset.Fields(j).Name
        Next
        
'        For Each c In Range(Range("K2"), Range("K2").End(xlDown))
'            c.Value = CInt(c.Value)
'        Next
       'Range(Range("K2"), Range("K2").End(xlDown)).NumberFormat = "0"
        'Range(Range("K2"), Range("K2").End(xlDown)).NumberFormat = "General"
        Range(Range("K2"), Range("K2").End(xlDown)).Value = Range(Range("K2"), Range("K2").End(xlDown)).Value
        maxtier = WorksheetFunction.Max(Range("K:K"))
        For j = 1 To maxtier
            Range("L1").Offset(0, j).Value = "Tier " & j
        Next
        
        'Roll up tiers
        '===================================================================================================================================================================
        Range("A2").Select
        ttlrws = Range("A1").End(xlDown).Row
        Do
            mtchcats = Application.CountIf(Range("A:A"), ActiveCell.Value)
            Set FirstCat = ActiveCell
            Application.StatusBar = "Transposing tiers: " & FirstCat.Row & " of " & ttlrws
            
            If Not mtchcats = 1 Then
                Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))
                
                'find active tiers to be replaced by future tiers and mark for deletion
                '----------------------------
                For Each c In AllCats
                    If Range("L" & c.Row).Value = "FUTURE" Then
                        For Each d In AllCats
                            If Not d.Row = c.Row And Range("K" & d.Row).Value = Range("K" & c.Row).Value And Range("F" & d.Row).Value = Range("F" & c.Row).Value Then d.Offset(0, 12 + maxtier).Value = 1
                        Next
                    End If
                Next
            
                'AllCats.Offset(0, 11).ClearContents 'clear future/active in column L
                prognmbr = 0
                For Each c In AllCats
                    If Not c.Offset(0, 12 + maxtier).Value = 1 Then
                        
                        If Not c.Row = FirstCat.Row And Not c.Offset(0, 5).Value = c.Offset(-1, 5).Value Then
                            prognmbr = c.Row - FirstCat.Row
                        'If ActiveCell.Offset(prognmbr, 0).Row = c.Row Then
                            Tier = Range("K" & c.Row).Value
                            FirstCat.Offset(prognmbr, 11 + Tier).Value = Range("J" & c.Row).Value
                            'If Not tier = 1 Then Range("K" & c.Row).ClearContents
                        Else
                            c.Offset(0, 12 + maxtier).Value = 1
                            FirstCat.Offset(prognmbr, 11 + Range("K" & c.Row).Value).Value = Range("J" & c.Row).Value
                        End If
                        'If Not ActiveCell.Offset(prognmbr, 9 + Range("K" & c.Row).Value).Value = 2 Then
                            'ActiveCell.Offset(prognmbr, 9 + Range("K" & c.Row).Value).Value
'                    ElseIf ActiveCell.Row = c.Row Then
'                        Range("J" & c.Row).ClearContents
                    End If
                Next
                'If Not prognmbr = mtchcats Then Range(ActiveCell.Offset(prognmbr + 1, 10), ActiveCell.Offset(mtchcats - prognmbr, 10)).ClearContents
                ActiveCell.Offset(0, 12 + maxtier).ClearContents
            Else
                'Range("L" & FirstCat.Row).ClearContents
                Tier = Range("K" & FirstCat.Row).Value
                FirstCat.Offset(0, 11 + Tier).Value = Range("J" & FirstCat.Row).Value
                'If Not tier = 1 Then Range("K" & FirstCat.Row).ClearContents
            End If
            
            FirstCat.Offset(mtchcats, 0).Select
        Loop Until Trim(ActiveCell.Value) = ""
        
        recset.Close
        On Error Resume Next
        Range("K:L").Delete Shift:=xlToLeft
        Range("A1").Offset(0, 10 + maxtier).EntireColumn.SpecialCells(xlCellTypeConstants).EntireRow.Select
        Application.StatusBar = "Deleting unused...Please wait"
        
        If Not SetupSwitch = 2 Then
            sortsht = connmbr
        Else
            sortsht = FUN_SuppName(i)
        End If
            
        Call FUN_Sort(sortsht & " Pricing", Range(Range("A2"), Range("A1").End(xlDown).Offset(0, 10 + maxtier)), Range("A1").Offset(0, 10 + maxtier), 2)
        Range(Range("A2").Offset(0, 10 + maxtier), Range("A2").Offset(0, 10 + maxtier).End(xlDown)).SpecialCells(xlCellTypeConstants).EntireRow.Delete Shift:=xlUp
        
'        itms = Application.CountIf(Range("A:A").Offset(0, 10 + maxtier), 1)
'        If itms < 10000 Then
'            Selection.Delete Shift:=xlUp
'        Else
'            Do Until Application.CountIf(Range("A:A").Offset(0, 10 + maxtier), 1) = 0
'                DoEvents
'                dltd = dltd + 5000
'                Application.StatusBar = "Deleting unused: " & dltd & " of " & itms
'                dltStrt = Range("A1").Offset(0, 10 + maxtier).End(xlDown).Row
'                Range("A" & dltStrt & ":A" & dltStrt + 5000).Offset(0, 10 + maxtier).SpecialCells(xlCellTypeConstants).EntireRow.Delete Shift:=xlUp
'            Loop
'        End If
        
        'Formatting
        '===================================================================================================================================================================
        Application.StatusBar = "Formatting...Please wait"
        
        'Remove zeros
        '----------------------------
        Range(Range("K2"), Range("A1").End(xlDown).Offset(0, 9 + maxtier)).Replace what:=0, replacement:="", lookat:=xlWhole
        On Error GoTo 0
        
        'Enter best price formula
        '----------------------------
        Range("J1").Value = "Tier Used"
        If SetupSwitch = 2 Then
            TierUsed = ConTblBKMRK.Offset(8, i).Address
            PFHdrRng = "$K$1:" & Range("J1").Offset(0, maxtier).Address
            PFTierRng = "K2:" & Range("J2").Offset(0, maxtier).Address(0, 0)
            Range("J2").Formula = "=IF(Index!" & TierUsed & "=""Best Price"",MIN(" & PFTierRng & "),OFFSET($J$1,ROW($J1),MATCH(Index!" & TierUsed & "," & PFHdrRng & ",0)))"
            If Not Trim(Range("J3").Value) = "" Then Range("J2").AutoFill Destination:=Range(Range("J2"), Range("A1").End(xlDown).Offset(0, 9))
        Else
            'Range(Range("J2"), Range("A1").End(xlDown).Offset(0, 9)).FormulaR1C1 = "=Min(RC11:RC" & 10 + maxtier & ")"
        End If

        'Create EA column
        '----------------------------
        Range("A1").Offset(0, 10 + maxtier).Value = "EA Price"
        Range(Range("A1").Offset(1, 10 + maxtier), Range("A1").End(xlDown).Offset(0, 10 + maxtier)).FormulaR1C1 = "=RC10/RC5"
        Range(Range("A1").Offset(1, 10 + maxtier), Range("A1").End(xlDown).Offset(0, 10 + maxtier)).NumberFormat = "$#,##0"
        
        'Create orgnl vals column
        '----------------------------
        Range("A1").Offset(0, 11 + maxtier).Value = "Original UOMs"
        Range(Range("E2"), Range("E1").End(xlDown)).Copy
        Range("A1").Offset(1, 11 + maxtier).PasteSpecial Paste:=xlPasteValues
        
        'Remove 0 price items
        '----------------------------
        'Call FUN_Sort(ActiveSheet.Name, Range(Range("A2"), Range("A1").End(xlDown).Offset(0, 11 + maxtier)), Range("J1"), 2)
        If Application.CountIf(Range("J:J"), 0) > 0 Then
            'Range(Range("J1").End(xlDown).Offset(-Application.CountIf(Range("J:J"), 0) + 1, 0), Range("J1").End(xlDown)).Select
            For Each c In Range(Range("J2"), Range("J1").End(xlDown))
                If c.Value = 0 Then
                    c.Offset(0, maxtier + 3).Value = c.Offset(0, -9).Value
                    c.Offset(0, -9).ClearContents
                End If
            Next
        End If

        'Add to keyword list
        '----------------------------
        On Error Resume Next
        If Not pfsingle = 1 Then 'And Not SetupSwitch = 2 Then
            Range(Range("C2"), Range("C1").End(xlDown)).Copy
            If Not pfsingle = 1 Then Sheets("pricefile keywords").Range("A1").End(xlDown).PasteSpecial Paste:=xlPasteValues
        End If
        On Error GoTo 0
        
        If SetupSwitch = 2 Then
        
            'add to contract table drop down
            '----------------------------
            With ConTblBKMRK.Offset(8, conpos).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='" & FUN_SuppName(i) & " Pricing'!J1:" & Range("J1").Offset(0, maxtier).Address(0, 0)
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = False
            End With
            ConTblBKMRK.Offset(8, conpos).Value = "Best Price"
            ConTblBKMRK.Offset(0, conpos).Value = PricingNm
            
            'Conditional formatting
            '-------------------------
            If Not QCFlg = True Then
                firsttier = ConTblBKMRK.Offset(9, i).Address(0, 0)
                Set tiersrng = Range(ConTblBKMRK.Offset(9, i), ConTblBKMRK.Offset(8 + maxtier, i))
                PFTierRng = "$K$2:" & Range("J2").Offset(0, maxtier).Address
                
                '[tbd] hardcode best tier based on sum
                MinTier = 1
                MinSum = WorksheetFunction.sum(Sheets(FUN_SuppName(i) & " Pricing").Range("K:K"))
                For j = 1 To maxtier - 1
                    CurrSum = WorksheetFunction.sum(Sheets(FUN_SuppName(i) & " Pricing").Range("K:K").Offset(0, j))
                    If CurrSum < MinSum Then
                        MinSum = CurrSum
                        MinTier = j + 1
                    End If
                Next
                
                With tiersrng
                     .FormatConditions.Delete
                     '.FormatConditions.Add Type:=xlExpression, Formula1:="=ROW(" & firsttier & ")=IF(" & TierUsed & "=""Best Price"",ROW(OFFSET(" & TierUsed & ",VALUE(SUBSTITUTE(OFFSET('" & FUN_SuppName(i) & " Pricing'!$J$1,0,MATCH(MIN('" & FUN_SuppName(i) & " Pricing'!" & PFTierRng & "),'" & FUN_SuppName(i) & " Pricing'!" & PFTierRng & ",0)),""Tier"","""")),0)),ROW(OFFSET(" & TierUsed & ",VALUE(SUBSTITUTE(" & TierUsed & ",""Tier"","""")),0)))"
                     .FormatConditions.Add Type:=xlExpression, Formula1:="=ROW(" & firsttier & ")=IF(" & TierUsed & "=""Best Price"",ROW(OFFSET(" & TierUsed & "," & MinTier & ",0)),ROW(OFFSET(" & TierUsed & ",VALUE(SUBSTITUTE(" & TierUsed & ",""Tier"","""")),0)))"
                     .FormatConditions(1).Interior.ColorIndex = 15
                End With
                With tiersrng.FormatConditions(1).Interior
                    .Pattern = xlPatternLinearGradient
                    .Gradient.Degree = 90
                    .Gradient.ColorStops.Clear
                End With
                With tiersrng.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                With tiersrng.FormatConditions(1).Interior.Gradient.ColorStops.Add(0.5)
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0
                End With
                With tiersrng.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End If
            
            'reattach Supplier headers on NC and conv tabs
            '-------------------------------------------------
            suppnmAddress = "Index!" & ConTblBKMRK.Offset(0, i).Address
            pftab = "'" & FUN_SuppName(i) & " Pricing'!"
            NonConBKMRK.Offset((MbrNMBR + 8) * (i - 1) - 2, 0).Formula = "=CONCATENATE(" & suppnmAddress & ","" "",Index!" & ConTblBKMRK.Offset(1, i).Address & ","" - "",IF(Index!" & ConTblBKMRK.Offset(8, i).Address & "=""Best Price"",OFFSET(" & pftab & "$J$1,0,MATCH(MIN(" & pftab & PFTierRng & ")," & pftab & PFTierRng & ",0)),Index!" & ConTblBKMRK.Offset(8, i).Address & "))"
            ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1) - 2, 0).Formula = "=CONCATENATE(" & suppnmAddress & ","" "",Index!" & ConTblBKMRK.Offset(1, i).Address & ","" - "",IF(Index!" & ConTblBKMRK.Offset(8, i).Address & "=""Best Price"",OFFSET(" & pftab & "$J$1,0,MATCH(MIN(" & pftab & PFTierRng & ")," & pftab & PFTierRng & ",0)),Index!" & ConTblBKMRK.Offset(8, i).Address & "))"
            
        End If

        'Clean up
        '----------------------------
'        Range("A:A").TextToColumns Destination:=Range("A:A"), DataType:=xlDelimited, _
'            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'            :=Array(1, 1), TrailingMinusNumbers:=True
        Range("J1").Offset(0, maxtier + 3).Value = "Removed"
        Range(Range("J2"), Range("J1").End(xlDown)).NumberFormat = "$#,##0"
        Range("B:B,E:E").Columns.ColumnWidth = 6
        Range("D:D").EntireColumn.Hidden = True
        Range("G:G").EntireColumn.Hidden = True
        Application.CutCopyMode = False
        Range("A1").Select
        
nxtCon: Next

    On Error Resume Next
    If Not pfsingle = 1 Then Sheets("Pricefile keywords").Range("A1:A2").Delete Shift:=xlUp
    If Not MainCall = 1 Then Application.StatusBar = False
    Set conn = Nothing
    Set recset = Nothing
    On Error GoTo 0
    
    
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_StandAlone:
Resume presTemp
presTemp:
    On Error GoTo 0
    Sheets.Add After:=Sheets(Sheets.Count)
    On Error Resume Next
    ActiveSheet.Name = connmbr & " Pricing"
    On Error GoTo 0
    If pfsingle = 1 Then i = 1
GoTo NotPres

ERR_LongName:
    PricingNm = Fun_ShortName(PricingNm)
    longerr = longerr + 1
    If longerr = 2 Then
        PricingNm = PricingNm & "2"
    ElseIf longerr = 3 Then
        PricingNm = Left(PricingNm, Len(PricingNm) - 1) & "2"
    ElseIf longerr = 4 Then
        PricingNm = i
    End If
Resume

errhndlNORECSET:
Range("A1").Value = "No Pricing Found"
Set recset = Nothing
Resume nxtCon


End Sub
Function Fun_ShortName(NmStr) As String

Dim parseName As String
parseName = NmStr
parseName = FUN_AlphaNumeric(parseName, 1)
Do Until Len(parseName) < 16
    parseName = Mid(parseName, 1, InStrRev(parseName, " ") - 1)
Loop
Fun_ShortName = parseName


End Function
Sub Create_Xref()

    'Setup Wb
    '-------------------------
    Call FUN_TestForSheet("DATxref")
    
    For Each sht In ActiveWorkbook.Sheets
        If ActiveSheet.Name = "DATxref" And sht.Visible = True Then sht.Select
        sht.Cells.UnMerge
        Cells.Interior.ColorIndex = 0
    Next
    
    'Bring up xref Ctrl
    '-------------------------
    CreateXrefForm.Show (False)

    
End Sub
Sub ExtractSupplierXref()
    
    'get form data and set variables
    '---------------------------------
    Set CreateXrefWB = ActiveWorkbook
    currsht = Replace(Left(CreateXrefForm.FormCatnum.Text, InStr(CreateXrefForm.FormCatnum.Text, "!") - 1), "'", "")
    Sheets(currsht).Select
    MftrNm = Trim(CreateXrefForm.FormName.Text)
    Set CatnumCol = Range(Mid(CreateXrefForm.FormCatnum.Text, InStr(CreateXrefForm.FormCatnum.Text, "!") + 1, Len(CreateXrefForm.FormCatnum.Text)))
    CatnumCol.Interior.Color = 16711935
    If Not CreateXrefForm.FormDesc.Text = "" Then
        Set DescCol = Range(Mid(CreateXrefForm.FormDesc.Text, InStr(CreateXrefForm.FormDesc.Text, "!") + 1, Len(CreateXrefForm.FormDesc.Text)))
        DescCol.Interior.Color = 16711935
    End If
    
    'enter data on DATxref tab
    '---------------------------------
    On Error GoTo NewMftrCol
    SuppCol = Sheets("DATxref").Rows("1:1").Find(what:=MftrNm, lookat:=xlWhole).Column
    On Error GoTo 0
    Set suppstrt = Sheets("DATxref").Cells(FUN_lastrow(SuppCol, "DATxref") + 1, SuppCol)
    Range(suppstrt, suppstrt.Offset(CatnumCol.Count - 1, 0)).Value = CatnumCol.Value
    If Not CreateXrefForm.FormDesc.Text = "" Then Range(suppstrt.Offset(0, 1), suppstrt.Offset(DescCol.Count - 1, 1)).Value = DescCol.Value

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'    'Set desc = Range("A1").Offset(0, DescAdd.Column - 1)
'    Sheets("DATxref").Range("A1").Value = mfgname
'    Range(col, col.Offset(FUN_lastrow(col.Column), 0)).Copy
'    Sheets("datxref").Range("A2").PasteSpecial xlPasteValues
'    If Not desc = "" Then
'        Range(desc, desc.Offset(FUN_lastrow(desc.Column), 0)).Copy
'        Sheets("datxref").Range("B2").PasteSpecial xlPasteValues
'    End If
'    col.EntireColumn.Interior.Color = 16711935
'    colCnt = 1
'    MaxRow = 1
'
'    'indefinite iterations
'    '------------------------
'    Do
'        On Error GoTo errhndlDone
'        Set coladd = Application.InputBox("Please select next supplier.  If none, press cancel.", Type:=8)
'        On Error GoTo 0
'        If Not coladd.Parent.Name = currsht Then
'            currsht = coladd.Parent.Name
'            GoTo nxtTab
'        End If
'        Set col = Range("A1").Offset(0, coladd.Column - 1)
'        mfgname = Trim(Application.InputBox("please enter the manufacturer name for this column.", Type:=2))
'        On Error GoTo errhndlNoDesc
'        Set DescAdd = Application.InputBox("please select associated description column", Type:=8)
'        Set desc = Range("A1").Offset(0, DescAdd.Column - 1)
'NoDesc: On Error GoTo newMfgCol
'        mfgcoladd = Sheets("datxref").Rows("1:1").Find(what:=mfgname, lookat:=xlWhole).Address
'        Set mfgcol = Sheets("DATxref").Range(mfgcoladd)
'        On Error GoTo 0
'        Range(col, col.Offset(FUN_lastrow(col.Column), 0)).Copy
'        mfgcol.Offset(MaxRow, 0).PasteSpecial xlPasteValues
'        If Not DescAdd Is Nothing Then
'            Range(desc, desc.Offset(FUN_lastrow(desc.Column), 0)).Copy
'            mfgcol.Offset(MaxRow, 1).PasteSpecial xlPasteValues
'        End If
'        col.EntireColumn.Interior.Color = 16711935
'        colCnt = colCnt + 1
'    Loop Until col = "xxindefinite placeholderxx"
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
NewMftrCol:
    If Trim(Sheets("DATxref").Range("A1").Value) = "" Then
        SuppCol = 1
        Sheets("DATxref").Range("A1").Value = MftrNm
        Sheets("DATxref").Range("A1").Font.Bold = True
        Sheets("DATxref").Range("A1").Interior.ColorIndex = 15
    Else
        SuppCol = Sheets("DATxref").Cells(1, FUN_lastcol(1, "DATxref") + 1).Column
        If ((SuppCol Mod 2) = 0) Then SuppCol = SuppCol + 1
        Sheets("DATxref").Cells(1, SuppCol).Value = MftrNm
        Sheets("DATxref").Cells(1, SuppCol).Font.Bold = True
        Sheets("DATxref").Cells(1, SuppCol).Interior.ColorIndex = 15
    End If
Resume Next


End Sub
Sub EndCreateXref()


    CreateXrefWB.Activate
    Call ImportXrefFromDB  '>>>>>>>>>>
    Application.DisplayAlerts = False
    If InStr(CreateXrefWB.Name, "CoreXref") > 0 Then
        CreateXrefWB.Save
    Else
        CreateXrefWB.SaveAs ZeusPATH & "\CoreXref(" & FileName_PSC & ")", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    Application.DisplayAlerts = True
    CreateXrefWB.Close


End Sub
Sub ImportXrefFromDB()

    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset
    
    On Error Resume Next
    If Not ZeusForm.asscContracts.ListCount = 0 Then
        Application.StatusBar = "Connecting to XrefDB..."
        conn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & XrefdbPATH

        For i = 0 To suppNMBR - 1
            connmbr = ZeusForm.asscContracts.List(i)
            Application.StatusBar = "Searching for " & connmbr & " Xref in Xref DB"
            
            'check if already present
            '--------------------------
            For Each sht In ActiveWorkbook.Sheets
                If sht.Name = "XrefDB_" & connmbr Then GoTo NxtSupp
            Next
            
            'Search for Xref
            '--------------------------
            Sheets.Add
            ActiveSheet.Name = "XrefDB_" & connmbr
            sqlstr = "" & _
                "SELECT [Xref Table].[Member Catalog Number], [Xref Table].[Contract Catalog Number], [Xref Table].[Cross Reference Supplier], [Xref Table].[Member Item Description], [Xref Table].[Contract Item Description], [Xref Table].Source, [Xref Table].Contract, [Xref Table].Date " & _
                "FROM [Xref Table] " & _
                "WHERE [Xref Table].Contract = '" & connmbr & "' AND NOT [Xref Table].[Member Catalog Number] IS NULL And NOT [Xref Table].[Contract Catalog Number] IS NULL"
        
            On Error GoTo ERR_NoXref
            recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

            If recset.RecordCount = 0 Then
                Range("A1").Value = "No xref found"
            Else
                Range("A1").CopyFromRecordset recset
            End If
NxtSupp:
            On Error Resume Next
        Next
    End If
        
Application.StatusBar = False
Exit Sub
':::::::::::::::::::::::::::::::::::::::
ERR_NoXref:
Range("A1").Value = "No xref found"
Resume NxtSupp

       
    
End Sub
Sub Import_Scopeguide()

    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    
    
    Call FUN_TestForSheet("Scopeguide")
    Cells.Clear

    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & ScopeguidePATH
    
    'Set adoRecSet = New ADODB.Recordset
    
    'Test if PSC exists
    '---------------------------------------
    sqlstr = "SELECT PSC FROM PSCs WHERE PSC = '" & PSCVar & "'"
    On Error GoTo errhndlNOSCOPE
    adoRecSet.Open sqlstr, ActiveConnection:=connDB, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    adoRecSet.MoveFirst
    Set adoRecSet = Nothing
    On Error GoTo 0
    
    'Find in scope
    '------------------------
    SELECTstr = "SELECT [Pim Key], [Standard Manufacturer], [Standard Catalog Number], [PSC], [In or Out]"
    FROMstr = " FROM [Main Table]"
    WHEREstr = " WHERE PSC = '" & PSCVar & "' AND [In or Out] = 'In Scope'"
    sqlstr = SELECTstr & FROMstr & WHEREstr
    
    On Error GoTo errhndlNOIN
    adoRecSet.Open sqlstr, ActiveConnection:=connDB, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    adoRecSet.MoveFirst
    On Error GoTo 0
    
    Range("A3").CopyFromRecordset adoRecSet
    Range(Range("A3"), Range("A3").End(xlDown).Offset(0, 4)).Interior.Color = 10213316
    Range(Range("A3"), Range("A3").End(xlDown).Offset(0, 4)).Borders.LineStyle = xlContinuous
    Set adoRecSet = Nothing

NoIn:
    'Find out of scope
    '------------------------
    SELECTstr = "SELECT [Pim Key], [Standard Manufacturer], [Standard Catalog Number], [PSC], [In or Out]"
    FROMstr = " FROM [Main Table]"
    WHEREstr = " WHERE PSC = '" & PSCVar & "' AND [In or Out] = 'Out of Scope'"
    sqlstr = SELECTstr & FROMstr & WHEREstr
    
    On Error GoTo errhndlNOOUT
    adoRecSet.Open sqlstr, ActiveConnection:=connDB, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    adoRecSet.MoveFirst
    On Error GoTo 0
    
    Range("G3").CopyFromRecordset adoRecSet
    Set adoRecSet = Nothing
    Set connDB = Nothing
    
    Range(Range("G3"), Range("G3").End(xlDown).Offset(0, 4)).Interior.Color = 12040422
    Range(Range("G3"), Range("G3").End(xlDown).Offset(0, 4)).Borders.LineStyle = xlContinuous
    
EndClean:
    If Not Range("C1").Value = "In Scope" Then Call ScopeGuide_Template
    

Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNOSCOPE:
Sheets("Scopeguide").Range("A1").Value = "No Scopeguide found for this PSC"
Sheets("Scopeguide").Range("A1").Interior.ColorIndex = 3
Sheets("Scopeguide").Range("A:A").Columns.AutoFit
Exit Sub

errhndlNOIN:
On Error GoTo 0
Resume NoIn

errhndlNOOUT:
Resume EndClean


End Sub
Sub Import_AdminFees()

    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset
    Dim conpos As Integer
    
    DoEvents
    Application.StatusBar = "Importing admin fees..."
    
    'Setup
    '====================================================================================================================
    Call FUN_TestForSheet("Admin Fees")
    Cells.Clear
    If Not Range("A1").Value = "Product ID" Then Call AdminFees_Template
    
    'Import
    '====================================================================================================================
    For i = 1 To suppNMBR
        conlist = conlist & "'" & ZeusForm.asscContracts.List(i - 1) & "', "
    Next
    conlist = Left(conlist, Len(conlist) - 2)
    
    SELECTstr = "SELECT CONTRACT_NUMBER, CONTRACT_NAME, REVENUE_FEE_DETAIL, REVENUE_PCT"
    FROMstr = " FROM OCSDW_CONTRACT"
    WHEREstr = " WHERE CONTRACT_NUMBER IN (" & conlist & ") ORDER BY contract_Number"
    
    sqlstr = SELECTstr & FROMstr & WHEREstr

    On Error GoTo errhndlNORECSET
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    
    Range("I2").CopyFromRecordset recset

    Set conn = Nothing
    Set recset = Nothing
    
    'Assign Fees
    '====================================================================================================================
    If SetupSwitch = 2 Then
        For i = 1 To suppNMBR
        
            'Get pricefile items
            '-------------------------------
            connmbr = ZeusForm.asscContracts.List(i - 1)
'            On Error GoTo errhndlNotPres
'            ConPos = Sheets("Impact Summary").Range(Bkmrk).Offset(1, 0).EntireRow.Find(what:=ConNmbr).Column - 1
'            On Error GoTo 0
'            Call FUN_SupplierSheet(ConPos, "Pricing")
'            If Trim(Range("A2").Value) = "" Then GoTo NotPres
            Set suppsht = Sheets(FUN_SuppName(i) & " Pricing")
            If suppsht.Range("A1").Value = "No Pricing Found" Then GoTo NotPres
            If suppsht.AutoFilterMode = True Then suppsht.AutoFilterMode = False
            Range(suppsht.Range("A2"), suppsht.Range("F1").End(xlDown)).Copy
            If Trim(Sheets("admin fees").Range("F2").Value) = "" Then
                LastRow = 2
            Else
                LastRow = Sheets("Admin Fees").Range("F1").End(xlDown).Row + 1
            End If
            Sheets("Admin Fees").Range("A" & LastRow).PasteSpecial Paste:=xlPasteValues
            
            'Find fee
            '-------------------------------
            If Not Trim(Sheets("Admin Fees").Range("L1").Offset(i, 0).Value) = 0 Then
                FeeVal = Sheets("Admin Fees").Range("L1").Offset(i, 0).Value
            Else
                On Error Resume Next
                feeStr = Sheets("Admin Fees").Range("K1").Offset(i, 0).Value
                If Not Mid(feeStr, InStr(feeStr, "%") - 2, 1) = "." Then
                    FeeVal = Mid(feeStr, InStr(feeStr, "%") - 1, 1) / 100
                Else
                    FeeVal = Mid(feeStr, InStr(feeStr, "%") - 3, 3) / 100
                End If
                On Error GoTo 0
                If InStr(LCase(feeStr), "novaplus") > 0 Or InStr(LCase(feeStr), "of all private label sales") > 0 Then
                    If InStr(LCase(feeStr), "novaplus") > 0 Then
                        novakey = "novaplus"
                    Else
                        novakey = "of all private label sales"
                    End If
                    On Error Resume Next
                    If Not Mid(feeStr, InStr(LCase(feeStr), novakey) - 3, 1) = "." Then
                        FeeVal = FeeVal + (Mid(feeStr, InStr(LCase(feeStr), novakey) - 3, 1) / 100)
                    Else
                        FeeVal = FeeVal + (Mid(feeStr, InStr(LCase(feeStr), novakey) - 5, 3) / 100)
                    End If
                    On Error GoTo 0
                End If
            End If
            Range(Sheets("Admin Fees").Range("G" & LastRow), Sheets("Admin Fees").Range("B1").End(xlDown).Offset(0, 5)).Value = FeeVal
            Range(Sheets("Admin Fees").Range("G" & LastRow), Sheets("Admin Fees").Range("B1").End(xlDown).Offset(0, 5)).NumberFormat = "0.00%"
NotPres:
        Next
    End If
    
    Sheets("Admin Fees").Select
    Range("I:K").Columns.AutoFit
    Range("A2").Select
    Range("L:L").ClearContents
    Range("H1:H" & Range("B1").End(xlDown).Row).Interior.ColorIndex = 1
    
    
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNORECSET:
Range("H1").Value = "No admin fees found."
Range("H1").Interior.ColorIndex = 3
Range("H1").VerticalAlignment = xlTop
Set recset = Nothing
Set conn = Nothing
Exit Sub

errhndlNotPres:
Range("H1").Value = Range("H1").Value & " No pricing tab found for " & connmbr
Range("H1").Interior.ColorIndex = 3
Range("H1").VerticalAlignment = xlTop
On Error GoTo 0
Resume NotPres



End Sub
Sub Import_Benchmarking()

Dim conn As New ADODB.Connection
Dim recset As New ADODB.Recordset

    'Setup
    '====================================================================================================================
    Application.StatusBar = "Importing Benchmarking..."
    Call FUN_TestForSheet("Best Market Price")
    Cells.Clear
    Call Benchmark_Template

'    'find most recently modified/created file with "benchmark product" & ".xlsx"
'    '=================================================================================================
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'    Set objFolder = objFSO.GetFolder(BenchPATH)
'    dteFile = DateSerial(1900, 1, 1)
'    For Each ofile In objFolder.Files
'        If ofile.DateCreated > dteFile And InStr(LCase(ofile.Name), LCase("Benchmark")) > 0 And InStr(ofile.Name, ".accdb") > 0 Then         '<<Use for date created
'        'If oFile.DateLastModified > dteFile Then   '<<Use for date modified
'            'Debug.Print oFile.Name
'            If Not InStr(ofile.Name, "~$") > 0 Then
'                'dteFile = oFile.DateLastModified   '<<Use for date modified
'                dteFile = ofile.DateCreated         '<<Use for date created
'                CheckName = ofile.Name
'                BenchFilePATH = ofile.Path
'            End If
'        End If
'    Next ofile
    
    
    'retrieve query
    '------------------------
'    Set objFSO = CreateObject("Scripting.FileSystemObject")                    '<-------VUN
'    BenchmarkQueryPATH = FUN_ConvTags(AdminConfigStr, "BenchmarkQueryPATH")
'    Set fileObj = objFSO.GetFile(BenchmarkQueryPATH)
'    orgnlstr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
    
                                                                                '<-------Take this out for VUN
    orgnlstr = "" & _
                "SELECT " & _
                    "part_number, " & _
                    "pim_key, " & _
                    "Pct_tile_10, " & _
                    "Pct_tile_25, " & _
                    "Pct_tile_50, " & _
                    "sample_size, " & _
                    "contract_number, " & _
                    "description, " & _
                    "product_spend_category, " & _
                    "uom, " & _
                    "company_name " & _
                "FROM [Data] as b " & _
                "WHERE b.product_spend_category = ? or b.contract_number in (?) or b.part_number in (?)"
                
'                    orgnlstr = "" & _
'                "SELECT " & _
'                    "part_number, " & _
'                    "pim_key, " & _
'                    "Pct_tile_10, " & _
'                    "Pct_tile_25, " & _
'                    "Pct_tile_50, " & _
'                    "sample_size, " & _
'                    "contract_number, " & _
'                    "description, " & _
'                    "product_spend_category, " & _
'                    "uom, " & _
'                    "company_name " & _
'                "FROM [Data] b " & _
'                "WHERE b.product_spend_category = ? or b.contract_number in (?) or b.part_number in (?)"
'
'                                    orgnlstr = "" & _
'                "SELECT " & _
'                    "part_number, " & _
'                    "pim_key, " & _
'                    "Pct_tile_10, " & _
'                    "Pct_tile_25, " & _
'                    "Pct_tile_50, " & _
'                    "sample_size, " & _
'                    "contract_number, " & _
'                    "description, " & _
'                    "product_spend_category, " & _
'                    "uom, " & _
'                    "company_name " & _
'                "FROM [Data$]  " & _
'                "WHERE [Data$].product_spend_category = ? or [Data$].contract_number in (?) or [Data$].part_number in (?)"
 
                '"FROM [Sheet1$] as b "  '<--Excel
 
    sqlstr = orgnlstr

    'insert PSC
    '------------------------
    If Not PSCVar = "" Then
        sqlstr = Replace(sqlstr, "b.product_spend_category = ?", "b.product_spend_category = '" & PSCVar & "'")
    Else
        sqlstr = Replace(sqlstr, "b.product_spend_category = ? or", "")
    End If

'    'Insert PSC
'    '------------------------
'    If Not PSCVar = "" Then
'        SQLstr = Replace(SQLstr, "[Data$].product_spend_category = ?", "[Data$].product_spend_category = '" & PSCVar & "'")
'    Else
'        SQLstr = Replace(SQLstr, "[Data$].product_spend_category = ? or", "")
'    End If
    
    'insert contracts
    '------------------------
    If suppNMBR > 0 Then
        For i = 0 To suppNMBR - 1
            Constr = Constr & "'" & ZeusForm.asscContracts.List(i) & "', "
        Next
        Constr = Left(Constr, Len(Constr) - 2)
        'SQLstr = Replace(SQLstr, "e.contract_number in (?)", "e.contract_number in (" & constr & ")")  '<-------VUN
        sqlstr = Replace(sqlstr, "b.contract_number in (?)", "b.contract_number in (" & Constr & ")")
    Else
        'SQLstr = Replace(SQLstr, "e.contract_number in (?) or ", "")                                   '<-------VUN
        sqlstr = Replace(sqlstr, "b.contract_number in (?) or ", "")
    End If

'    'insert contracts
'    '------------------------
'    If suppNMBR > 0 Then
'        For i = 0 To suppNMBR - 1
'            Constr = Constr & "'" & ZeusForm.asscContracts.List(i) & "', "
'        Next
'        Constr = Left(Constr, Len(Constr) - 2)
'        'SQLstr = Replace(SQLstr, "e.contract_number in (?)", "e.contract_number in (" & constr & ")")  '<-------VUN
'        SQLstr = Replace(SQLstr, "[Data$].contract_number in (?)", "[Data$].contract_number in (" & Constr & ")")
'    Else
'        'SQLstr = Replace(SQLstr, "e.contract_number in (?) or ", "")                                   '<-------VUN
'        SQLstr = Replace(SQLstr, "[Data$].contract_number in (?) or ", "")
'    End If
    
    
    'insert items if <1000
    '------------------------
    If SetupSwitch = 2 Then
        itms = Application.CountA(Sheets("line item data").Range("A:A")) - 3
        If itms > 0 And itms < 1000 Then
            For i = 1 To itms
                ItmStr = ItmStr & "'" & Sheets("line item data").Range("X4").Offset(i, 0).Value & "', "
            Next
            ItmStr = Left(ItmStr, Len(ItmStr) - 2)
            sqlstr = Replace(sqlstr, "or b.part_number in (?)", "or b.part_number in (" & ItmStr & ")")
            Application.StatusBar = "Importing Benchmark Data for PSC, Contracts, and Items...Please Wait"
        Else
            sqlstr = Replace(sqlstr, "or b.part_number in (?)", "")
            Application.StatusBar = "Importing Benchmark Data for PSC and Contracts...Please Wait"
        End If
    Else
        sqlstr = Replace(sqlstr, "or b.part_number in (?)", "")
    End If

'    'insert items if <1000
'    '------------------------
'    If SetupSwitch = 2 Then
'        itms = Application.CountA(Sheets("line item data").Range("A:A")) - 3
'        If itms > 0 And itms < 1000 Then
'            For i = 1 To itms
'                ItmStr = ItmStr & "'" & Sheets("line item data").Range("X4").Offset(i, 0).Value & "', "
'            Next
'            ItmStr = Left(ItmStr, Len(ItmStr) - 2)
'            sqlstr = Replace(sqlstr, "or [Data$].part_number in (?)", "or [Data$].part_number in (" & ItmStr & ")")
'        Else
'            sqlstr = Replace(sqlstr, "or [Data$].part_number in (?)", "")
'        End If
'    Else
'        sqlstr = Replace(sqlstr, "or [Data$].part_number in (?)", "")
'    End If

    'return results
    '------------------------
    'conn.Open "Provider=SQLNCLI11;Server=dbvhadmprod;Database=VUN;Uid=bforrest;pwd=Lovepeople16"       '<-------VUN alt
    'conn.Open "Provider=SQLNCLI11;Server=dbvhadmprod;Database=VUN;Trusted_Connection=yes"              '<-------VUN
    'conn.Open "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & BenchFilePATH   '<-------Excel
    'conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FUN_ConvTags(AdminconfigStr, "LocalBenchPath") & "\Benchmarking.xlsx" & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"


    On Error GoTo ERR_NoLocal
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FUN_ConvTags(AdminconfigStr, "LocalBenchPath") & "\Benchmarking.accdb"
    On Error GoTo 0
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    Range("A2").CopyFromRecordset recset
    
    'search for items if > 1000
    '------------------------
    If itms > 1000 Then
        'orgnlstr = Replace(SQLstr, "or e.contract_number in (" & constr & ")", "and not e.contract_number in (" & constr & ")")    '<-------VUN
        orgnlstr = Replace(sqlstr, "or b.contract_number in (" & Constr & ")", "and not b.contract_number in (" & Constr & ")")
        orgnlstr = Replace(orgnlstr, "(b.product_spend_category = '" & PSCVar & "'", "not b.product_spend_category = '" & PSCVar & "'")
        For i = 0 To Round(itms / 1000 + 0.5) - 1
            recset.Close
            sqlstr = orgnlstr
            ItmStr = ""
            itms = Application.CountA(Sheets("line item data").Range("A:A")) - 3 - (i * 1000)
            If itms > 1000 Then itms = 1000
            For j = 1 To itms
                ItmCnt = ItmCnt + 1
                ItmStr = ItmStr & "'" & Sheets("line item data").Range("X4").Offset(i * 1000 + j, 0).Value & "', "  '<--to convert to number bc Bench data doesn't use leading zeros use: Val()
            Next
            ItmStr = Left(ItmStr, Len(ItmStr) - 2)
            sqlstr = Replace(sqlstr, ")group by", "and b.part_number in (" & ItmStr & ")group by")
            Application.StatusBar = "Importing Benchmark Data For Items: " & ItmCnt - itms & " to " & ItmCnt & "...Please Wait"
            recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
            Range("A1").End(xlDown).Offset(1, 0).CopyFromRecordset recset
        Next
    End If
    
    Set recset = Nothing
    Set conn = Nothing

    Range(Range("O2"), Range("A1").End(xlDown).Offset(0, 14)).Value = Range(Range("K2"), Range("K1").End(xlDown)).Value
    Range(Range("K2"), Range("K1").End(xlDown)).ClearContents
    
    'conver to number
    '------------------------------
'    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
'        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("A:A").NumberFormat = "General"
    
    
    'Bring in UOMs
    '------------------------------
    If SetupSwitch = 2 Then Call BMP_UOM_Lookup  '>>>>>>>>>>

    'fill in L:N formulas
    '------------------------------
    Range("L2").Formula = "=C2/K2"
    Range("M2").Formula = "=D2/K2"
    Range("N2").Formula = "=E2/K2"
    If Not IsEmpty(Range("A3")) Then Range("L2:N2").AutoFill Destination:=Range("L2:N" & Range("A1").End(xlDown).Row)
    Range("L:N").NumberFormat = "$#,##0"
    
    'finalize
    '-----------------------------
    Range("K:K").Copy
    Range("K:K").PasteSpecial Paste:=xlPasteValues
    Range("L:N").Calculate
    
    Application.StatusBar = "Sorting...Please wait"
    ActiveWorkbook.Worksheets("Best Market Price").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Best Market Price").Sort.SortFields.Add Key _
        :=Range("I2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Best Market Price").Sort.SortFields.Add Key _
        :=Range("A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Best Market Price").Sort.SortFields.Add Key _
        :=Range("F2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Best Market Price").Sort
        .SetRange Range("A2:Z100000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("K:K").Copy
    Range("P1").Insert Shift:=xlToRight
    Range("P1").Value = "Original values"
    
    Application.StatusBar = False
    Application.CutCopyMode = False
    Range("A1").Select
    If SetupSwitch = 2 Then
        
        'Enter Bench Data Stats
        '----------------------------
        Sheets("notes").Range("Q1").Value = "BENCHMARKING DATA:"
        Sheets("notes").Range("Q1").Font.Bold = True
        Sheets("notes").Range("Q1").Font.Underline = True
        
        'PSC
        '----------------------------
        pscItms = Application.CountIf(Range("I:I"), PSCVar)
        Sheets("notes").Range("Q2").Value = pscItms & " items found for " & PSCVar
        
        'Contracts
        '----------------------------
        For i = 0 To suppNMBR - 1
            connmbr = ZeusForm.asscContracts.List(i)
            ConItms = Application.CountIf(Range("G:G"), connmbr)
            ttlConItms = ttlConItms + ConItms
            Sheets("notes").Range("Q1").End(xlDown).Offset(1, 0).Value = ConItms & " items found for " & connmbr
        Next
        
        'unbound itms
        '----------------------------
        UbItms = FUN_lastrow("A") - ttlConItms - pscItms
        Sheets("notes").Range("Q1").End(xlDown).Offset(1, 0).Value = UbItms & " Unbound items found"
        
        Sheets("line item data").Range("AG:AJ").Calculate
        Range(MSGraphBKMRK.Offset(MbrNMBR + 8, -1), MSGraphBKMRK.Offset(MbrNMBR * 2 + 8, 8)).Calculate
    
        'Format Size & Position of benchmark graph
        '----------------------------------
        If MbrNMBR < 35 Then
            hghtcells = MbrNMBR + 1
        Else
            hghtcells = 35
        End If

        Set RngToCover = Range(MSGraphBKMRK.Offset(MbrNMBR + 8, 10), MSGraphBKMRK.Offset(MbrNMBR + 8 + hghtcells, 10))
        Set ChtObj = Sheets("initiative Spend overview").ChartObjects(1)
        ChtObj.Height = RngToCover.Height ' resize
        ChtObj.Width = ChtObj.Height   ' resize
    End If


Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_NoLocal:
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FUN_ConvTags(AdminconfigStr, "BenchPath") & "\Benchmarking.accdb"
Resume Next



End Sub
Sub BMP_UOM_Lookup()

'insert UOMs in the BMP tab if there's a match in the pricefile
'*****************************************************
    
    Application.StatusBar = "Looking up BMP UOMs..."
    Set BMPstrt = Sheets("Best Market Price").Range("A1")
    BMPitms = Application.CountA(Range(BMPstrt, BMPstrt.End(xlDown))) - 1
    
    'Search Pricefiles
    '----------------------
    On Error GoTo ERR_notFnd
    For i = 1 To suppNMBR
        Set SrchRng = Sheets(FUN_SuppName(i) & " Pricing").Range("A:A")
        For j = 1 To BMPitms
            Application.StatusBar = "Searching for UOMs in Supplier" & i & " Pricing: " & j & " of " & BMPitms
            If BMPstrt.Offset(j, 10).Value = "" Then BMPstrt.Offset(j, 10).Value = SrchRng.Find(what:=BMPstrt.Offset(j, 0).Value, lookat:=xlWhole).Offset(0, 4).Value  'Or IsError(Sheets("Best Market Price").Range("K1").Offset(BMPloop, 0)) Then
notFnd:
        Next
    Next
    
    'Search line item data
    '----------------------
    On Error GoTo ERR_notFnd2
    For i = 1 To BMPitms
        DoEvents
        Application.StatusBar = "Searching for UOMs in line item data: " & i & " of " & BMPitms
        If BMPstrt.Offset(i, 10).Value = "" Then BMPstrt.Offset(i, 10).Value = Sheets("line item data").Range("X:X").Find(what:=BMPstrt.Offset(i, 0).Value, lookat:=xlWhole).Offset(0, 4).Value  'Or IsError(Sheets("Best Market Price").Range("K1").Offset(BMPloop, 0)) Then
notFnd2:
    Next
    
Exit Sub
'::::::::::::::::::::::::::::::::
ERR_notFnd:
Resume notFnd

ERR_notFnd2:
Resume notFnd2


End Sub
'Sub ManualHCO()
'
''insert UOMs in the BMP tab if there's a match on HCO
''*****************************************************
'
'    On Error Resume Next
'
'    Set BMPstrt = Sheets("Best Market Price").Range("A1")
'    BMPitms = Application.CountA(Range(BMPstrt, BMPstrt.End(xlDown))) - 1
'    For i = 1 To BMPitms
'        DoEvents
'        Application.StatusBar = "Searching for UOMs in line item data: " & i & " of " & BMPitms
'        If Sheets("Best Market Price").Range("K1").Offset(i, 0).Value = "" Then BMPstrt.Offset(i, 10).Value = Sheets("line item data").Range("X:X").Find(what:=BMPstrt.Offset(i, 0).Value, lookat:=xlWhole).Offset(0, 4).Value  'Or IsError(Sheets("Best Market Price").Range("K1").Offset(BMPloop, 0)) Then
'    Next
'
'Exit Sub
'':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'
'End Sub

'Sub ManualBENCH()
'
''insert UOMs in the BMP tab if there's a match on HCO
''*****************************************************
'
'    bmWB.ActiveSheet.AutoFilterMode = False
'    bmWB.Sheets(1).Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, ConsecutiveDelimiter:=False, Tab:=False, _
'        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'        :=Array(1, 1), TrailingMinusNumbers:=True
'    bmWB.Sheets(1).Range("A:A").NumberFormat = "General"
'
'    tmWB.Activate
'    If Trim(Sheets("Best Market Price").Range("A2").Value) = "" Then Sheets("Best Market Price").Range("A2:O2").Value = "x"
'    lastHCO = Sheets("Line Item Data").Range("A1").End(xlDown).Row
'    lastbm = bmWB.Sheets(1).Range("A1").End(xlDown).Row
'    bmstrt = Sheets("Best Market Price").Range("A1").End(xlDown).Row
'    Application.ScreenUpdating = False
'    strttm = Time
'    extraNumber = 0
'    Sheets("Line Item Data").Select
'    Range("N3").Select
'    Do
'        DoEvents
'        Application.StatusBar = "Raising Benchmark Rate: " & ActiveCell.Row & " of " & lastHCO
'
'        'On Error GoTo errhndlALRDYFND
'        If Application.CountIf(Sheets("Best Market Price").Range("A:A"), ActiveCell.Value) > 0 Then GoTo NxtHCO '.Find(what:=ActiveCell.Value, lookat:=xlWhole).Value
'NoFND:  'On Error GoTo errhndlNOBMP
'        If Application.CountIf(Range(bmWB.Sheets(1).Range("A1"), bmWB.Sheets(1).Range("A" & lastbm)), ActiveCell.Value) > 0 Then
'            bmpfnd = Range(bmWB.Sheets(1).Range("A1"), bmWB.Sheets(1).Range("A" & lastbm)).Find(what:=ActiveCell.Value, lookat:=xlWhole).Address
'            Range(bmWB.Sheets(1).Range(bmpfnd), bmWB.Sheets(1).Range(bmpfnd).End(xlToRight)).Copy
'            Sheets("Best Market Price").Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
'            extranmbr = extranmbr + 1
'NxtBMP:     If Application.CountIf(Range(bmWB.Sheets(1).Range(bmpfnd).Offset(1, 0), bmWB.Sheets(1).Range("A1").End(xlDown)), ActiveCell.Value) > 0 Then
'                bmpfnd = Range(bmWB.Sheets(1).Range(bmpfnd).Offset(1, 0), bmWB.Sheets(1).Range("A" & lastbm)).Find(what:=ActiveCell.Value, lookat:=xlWhole).Address
'                Range(bmWB.Sheets(1).Range(bmpfnd), bmWB.Sheets(1).Range(bmpfnd).End(xlToRight)).Copy
'                Sheets("Best Market Price").Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
'                extranmbr = extranmbr + 1
'                GoTo NxtBMP
'            End If
'        End If
'NxtHCO: ActiveCell.Offset(Application.CountIf(Range(Range("N3"), Range("N" & lastHCO)), ActiveCell.Value), 0).Select
'    Loop Until IsEmpty(ActiveCell.Offset(1, 0))
'
'If Sheets("Best Market Price").Range("A2").Value = "x" Then Sheets("Best Market Price").Range("A2").EntireRow.Delete Shift:=xlUp
'
'endtm = Time
'If extranmbr = 0 Then
'    Sheets("Notes").Range("Q1").End(xlDown).Offset(1, 0).Value = "No unbound HCO catalog numbers found"
'Else
'    'c/p mfg name to source column
'    '--------------------------------
'    Sheets("Best market price").Select
'    Range("A:A").NumberFormat = "General"
'    If Not IsEmpty(Range("O2")) Then
'        Range(Range("O1").End(xlDown).Offset(1, -4), Range("A1").End(xlDown).Offset(1, 10)).Cut
'        Range("O1").End(xlDown).Offset(1, 0).Select
'    Else
'        Range(Range("K2"), Range("K1").End(xlDown)).Cut
'        Range("O2").Select
'    End If
'    ActiveSheet.Paste
'    Sheets("Notes").Range("Q1").End(xlDown).Offset(1, 0).Value = extranmbr & " unbound HCO catalog numbers found"
'End If
'
'
'Exit Sub
'':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'errhndlALRDYFND:
'On Error GoTo 0
'Resume NoFND
'
'errhndlNOBMP:
'On Error GoTo 0
'Resume NxtHCO
'
'
'
'End Sub
Sub Import_UNSPSC()

    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset
    
    'Setup
    '============================================================================================================
    If Not SetupSwitch = 2 Then
        Call FUN_TestForSheet("PRS & UNSPSC")
        Range("A2:B11").ClearContents
    End If

    'Build Query string
    '============================================================================================================
    'SqlSELECT = "SELECT TOP 10 psr.UNSPSC_TITLE, psr.UNSPSC_CODE, Count(psr.vendor_product_number) As Product_Count FROM"
    
    
    
'    SqlSELECT = "SELECT TOP 10 psr.UNSPSC_TITLE, psr.UNSPSC_CODE FROM"
'    SqlSubSELECT = " (SELECT DISTINCT con.CONTRACT_NUMBER, lip.vendor_product_number, lip.product_description, lip.program_name, lip.vendor_name, nmfa.UNSPSC_CODE, nmfa.UNSPSC_TITLE"
'    SqlSubFROM = " FROM OCSDW_CONTRACT con INNER Join (OCSDW_LINE_ITEM_PRICING_DETAIL lip LEFT JOIN NFMA_Product nmfa ON lip.product_key = nmfa.PRODUCT_KEY) ON con.CONTRACT_NUMBER = lip.contract_number"
'    SqlSubWHERE = " WHERE (con.CONTRACT_NUMBER Like '" & ZeusForm.asscContracts.List(0) & "'"
    For i = 2 To suppNMBR
        SQLAddCon = SQLAddCon & " OR con.CONTRACT_NUMBER Like '" & ZeusForm.asscContracts.List(i - 1) & "'"
    Next
'    SqlActFut = ") AND (lip.price_status_code='ACTIVE' Or lip.price_status_code='FUTURE')) AS psr"
'    SqlGB = " GROUP BY psr.UNSPSC_CODE, psr.UNSPSC_TITLE"
'    SqlHAVING = " HAVING psr.UNSPSC_CODE<>''"
'    SqlOB = " ORDER BY Count (psr.vendor_product_number) DESC;"
    
    'SQLstr = SqlSELECT & SqlSubSELECT & SqlSubFROM & SqlSubWHERE & SQLAddCon & SqlActFut & SqlGB & SqlHAVING & SqlOB
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    QueryPATH = FUN_ConvTags(AdminconfigStr, "UNSPSCQuery_Path")
    Set fileObj = objFSO.GetFile(QueryPATH)
    sqlstr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
    sqlstr = Replace(sqlstr, "!!CONTRACTID!!", ZeusForm.asscContracts.List(0))
    sqlstr = Replace(sqlstr, "!!ADDCONTRACT!!", SQLAddCon)
    
    'Connect & import
    '============================================================================================================
    conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    
    On Error GoTo errhndlNORECSET
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    recset.MoveFirst
    On Error GoTo 0
    
    If SetupSwitch = 2 Then
        Sheets("Index").Select
        Set UNSPSCstrt = Cells.Find(what:="UNSPSC", lookat:=xlPart).Offset(0, -3)
        For i = 1 To recset.RecordCount
            UNSPSCstrt.Offset(i, 0).Value = recset.Fields(0)
            UNSPSCstrt.Offset(i, 1).Value = recset.Fields(1)
            recset.MoveNext
        Next
        'ActiveCell.Offset(1, 0).CopyFromRecordset Recset
    Else
        Sheets("PRS & UNSPSC").Select
        Range("A2").CopyFromRecordset recset
    End If
    
    On Error Resume Next
    Set conn = Nothing
    Set recset = Nothing
    On Error GoTo 0
    
    'Clean and End
    '============================================================================================================
    If Not SetupSwitch = 2 Then
        Range(Range("C2"), Range("B2").End(xlDown).Offset(0, 1)).Interior.ColorIndex = 1
        Range(Range("A2"), Range("B2").End(xlDown)).Interior.Color = 14277081
        Range(Range("A2"), Range("B2").End(xlDown)).Borders.LineStyle = xlContinuous
        If Not Range("A1:Y1").Interior.ColorIndex = 16 Then Call PRSUNSPSC_Template
    End If


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNORECSET:
If SetupSwitch = 2 Then
    Sheets("index").Cells.Find(what:="UNSPSC", lookat:=xlPart).Offset(1, 0).Value = "No UNSPSC data found"
Else
    Range("A2").Value = "No UNSPSC data found"
End If
conn.Close
Set conn = Nothing
Exit Sub


End Sub
Sub Import_PRS(Optional AltPRS As Boolean)

'must get memIDs and hos names from sna engagement file, create table within vba and join to it ia memID within vba
    
    Dim conn As New ADODB.Connection
    Dim recset As New ADODB.Recordset
    Dim AnnFctr As Double


'build Query & Parameters
'=============================================================================================

    'find dates & annualization factor
    '----------------------
    If ZeusForm.NRSrch = False Then
        dateFld = "prs.SUPPL_DT_SKEY"
    Else
        dateFld = "prs.SUPPLIER_DATE"
    End If
    If Trim(ZeusForm.asscStartDate.Text) = "" Then
        strtdate = "to_number(to_char(add_months(trunc(mbr.spend_end_dt,'mm'),-11),'yyyymm')||'00')"
    Else
        strtdate = "'" & Format(ZeusForm.asscStartDate.Text, "yyyymmdd") & "'"
    End If
    If Trim(ZeusForm.asscEndDate.Text) = "" Then
        EndDate = "to_number(to_char(mbr.spend_end_dt,'yyyymm')||'99')"
    Else
        EndDate = "'" & Format(ZeusForm.asscEndDate.Text, "yyyymmdd") & "'"
    End If
    DateParam = " WHERE (" & dateFld & " BETWEEN " & strtdate & " AND " & EndDate & ")"
    
    If ZeusForm.AnnualizedChk.Value = True Then
        Numdays = DateDiff("d", ZeusForm.asscStartDate.Text, ZeusForm.asscEndDate.Text)
        AnnFctr = Format((365 / Numdays), "0.00")
    Else
        AnnFctr = 1
    End If
    
    'Find contracts
    '----------------------
    If ZeusForm.NRSrch = True Then
        confld = "prs.CONTRACT_ID"
    Else
        confld = "con.CONTR_NBR"
    End If
    
    If AltPRS = True Then
        AltPRSform.Show
        SQLConParam = confld & " LIKE '" & Trim(UCase(AltPRSform.Controls("AltPRScon1").Text)) & "%'"
        For i = 2 To suppNMBR
            SQLConParam = SQLConParam & " OR " & confld & " LIKE '" & Trim(UCase(AltPRSform.Controls("AltPRScon" & i).Text)) & "%'"
        Next
    Else
        SQLConParam = confld & " LIKE '" & UCase(ZeusForm.asscContracts.List(0)) & "%'"
        For i = 1 To suppNMBR - 1
            SQLConParam = SQLConParam & " OR " & confld & " LIKE '" & UCase(ZeusForm.asscContracts.List(i)) & "%'"
        Next
    End If
    SQLConParam = " AND (" & SQLConParam & ") "

'    'find MIDs
'    '----------------------
'    For i = 0 To Round(itms / 1000 + 0.5) - 1
'        for j = 1 to 1000
'            MIDstr = MIDstr & & ", "
'        next
'    Next
'    MIDstr = Left(MIDstr, Len(MIDstr) - 2)

    If ZeusForm.DSdefaultChk.Value = True Then

        If ZeusForm.NRSrch = True Then
            SELECTstr = "SELECT prs.member_id, (" & AnnFctr & " * SUM(prs.MEMBER_SALES_AMT)) AS SPEND_TTL, prs.CONTRACT_ID "
            FROMstr = "FROM PRS_MEMBER_CONTRACT_SALES_ALL prs " 'INNER JOIN MEMT1MEINQ ntwk ON ntwk.MEMID = prs.member_id"
            WHEREstr = DateParam & SQLConParam & SQLNetParam & SQLSysParam & SQLmbrParam & " GROUP BY prs.CONTRACT_ID, prs.member_id ORDER BY prs.CONTRACT_ID, prs.member_id"   'group by name?
        Else
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            QueryPATH = FUN_ConvTags(AdminconfigStr, "PRSQuery_Path")
            Set fileObj = objFSO.GetFile(QueryPATH)
            sqlstr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
            sqlstr = Replace(sqlstr, "!!ANNUALFACTOR!!", AnnFctr)
            sqlstr = Replace(sqlstr, "!!COMPANYCODE!!", CmpyCD)
            'SQLstr = Replace(SQLstr, "!!NETWORKID!!", NtwkIDArray(NetPos))
            sqlstr = Replace(sqlstr, "!!DATERANGE!!", DateParam)
            sqlstr = Replace(sqlstr, "!!CONTRACTID!!", SQLConParam)
            
            'Network ID
            '--------------------------
            If NetNm = "NUSPC" Then
                sqlstr = Replace(sqlstr, "!!NETWORKID!!", "net.agg_grp_sub_id = " & NtwkIDArray(NetPos) & " ")
            Else
                sqlstr = Replace(sqlstr, "!!NETWORKID!!", "net.agg_grp_id = " & NtwkIDArray(NetPos) & " ")
            End If
            
            'Participation
            '--------------------------
            If NetNm = "CHA" Then
                sqlstr = Replace(sqlstr, "!!PARTICIPATION!!", "(net.prtcptn_sts_cd = 'A' OR mbr.prime_typ_cd = 'C')")
            Else
                sqlstr = Replace(sqlstr, "!!PARTICIPATION!!", "net.prtcptn_sts_cd = 'A'")
            End If
'
'            SELECTstr = "" & _
'            "SELECT /*+ Full(prs,8) Full(con,8)*/ " & _
'                "mbr.MBR_ID, " & _
'                "(SUM(prs.MBR_SALES_AMT) * " & AnnFctr & ") AS PRS_SPEND, " & _
'                "con.CONTR_NBR "
'
'            FROMstr = "" & _
'            "FROM RDM_PRS_MBR_CONTR_SALES prs " & _
'            "INNER JOIN RDM_MBR_DIM mbr " & _
'            "ON mbr.MBR_SKEY = prs.MBR_SKEY " & _
'            "INNER JOIN rdm.rdm_mbr_agg_grp_dim net " & _
'            "ON mbr.mbr_skey = net.mbr_skey " & _
'            "AND net.agg_grp_sts_cd = 'A' " & _
'            "AND net.prtcptn_sts_cd = 'A' " & _
'            "AND net.agg_grp_id = " & NtwkIDArray(NetPos) & " " & _
'            "INNER JOIN RDM_CONTR_DIM con " & _
'            "ON con.CONTR_SKEY = prs.CONTR_SKEY "
'
'            WHEREstr = "" & _
'            DateParam & _
'            SQLConParam & _
'            "GROUP BY " & _
'            "mbr.MBR_ID, " & _
'            "con.CONTR_NBR " & _
'            "ORDER BY con.CONTR_NBR"
            
        End If

    Else
    
        'Find systems if not DS default
        '----------------------
        If ZeusForm.asscSystems.ListCount > 0 Then
            SQLSysParam = "ntwk.systemname = '" & ZeusForm.asscSystems.List(0) & "'"
            For i = 1 To ZeusForm.asscSystems.ListCount - 1
                SQLSysParam = SQLSysParam & " OR ntwk.systemname = '" & ZeusForm.asscSystems.List(i) & "'"
            Next
            SQLSysParam = " AND (" & SQLSysParam & ")"
        End If
        
        'Find individual members if not DS default
        '----------------------
        If ZeusForm.asscMembers.ListCount > 0 Then
            SQLmbrParam = "ntwk.name = '" & ZeusForm.asscMembers.List(0) & "'"
            For i = 1 To ZeusForm.asscMembers.ListCount - 1
                SQLmbrParam = SQLmbrParam & " OR ntwk.name = '" & ZeusForm.asscMembers.List(i) & "'"
            Next
            SQLmbrParam = " AND (" & SQLmbrParam & ")"
        End If

        SELECTstr = "SELECT (" & AnnFctr & " * SUM(prs.MEMBER_SALES_AMT)) AS SPEND_TTL, ntwk.name, ntwk.systemname, prs.CONTRACT_ID"
        FROMstr = " FROM PRS_MEMBER_CONTRACT_SALES_ALL prs INNER JOIN MEMT1MEINQ ntwk ON ntwk.MEMID = prs.member_id"
        WHEREstr = " WHERE (prs.SUPPLIER_DATE BETWEEN '" & strtdate & "' AND '" & EndDate & "')" & SQLConParam & SQLNetParam & SQLSysParam & SQLmbrParam & " GROUP BY prs.CONTRACT_ID, ntwk.systemname, ntwk.name ORDER BY prs.CONTRACT_ID, ntwk.name, ntwk.systemname"   'group by name?
        
    End If

    'import data
    '======================================================================================================================================
    Application.StatusBar = "Downloading PRS...Please wait"
    'SQLstr = SELECTstr & FROMstr & WHEREstr
    If ZeusForm.NRSrch = True Then
        On Error GoTo errhndlNORECSET
        connstr = "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
        conn.Open connstr
        recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    Else
        On Error GoTo errhndlReconnect
attempt2:
        If RDMconn = "" Then
            connstr = "Driver={Microsoft ODBC for Oracle};CONNECTSTRING=" & RDMConnStr
            RDMconn.Open connstr
        End If
        recset.Open sqlstr, ActiveConnection:=RDMconn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    End If
    On Error GoTo errhndlNORECSET
    
    'setup sheet to copy data from recset
    '----------------------
    If SetupSwitch = 2 Then
        Sheets("notes").Range("AD1:AF1").EntireColumn.ClearContents
        Sheets("notes").Range("AD1").Value = "UNASSIGNED PRS"
        Sheets("notes").Range("AD1").Font.Bold = True
    Else
        Call FUN_TestForSheet("PRS & UNSPSC")
        Range("D2:W500").Clear
    End If
    'If Recset.RecordCount = 0 Then GoTo errhndlNORECSET
    Call FUN_TestForSheet("xxCalculations")
    Cells.Clear
    Range("A1").CopyFromRecordset recset
    If Trim(Range("A1").Value) = "" Then GoTo errhndlNORECSET

    Set recset = Nothing
    On Error GoTo 0
    
    Application.StatusBar = "Assigning PRS..."
    
    'convert MIDs to member names
    '----------------------
'    If ZeusForm.NRsrch = True Then
'        Dim adoRecSet As New ADODB.Recordset
'        Dim connDB As New ADODB.Connection
'
'        Application.ScreenUpdating = False
'        'On Error GoTo errhndlNORECSET
'        connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & SupplyNetPATH & "Analytics\Analytical Tools\Pulling PRS Spend\PRS Generator.accdb"
'
'        strSQL = "SELECT * FROM [Network Members] WHERE [Network] = '" & NetNm & "'"
'        adoRecSet.Open strSQL, ActiveConnection:=connDB, CursorType:=adOpenStatic, LockType:=adLockOptimistic
'        adoRecSet.MoveFirst
'
'        Set fndrng = Range(Range("A1"), Range("A1").End(xlDown))
'        On Error Resume Next
'        For i = 1 To adoRecSet.RecordCount - 1
'            fndrng.Replace what:=adoRecSet.Fields(0), replacement:=adoRecSet.Fields(1), lookat:=xlWhole
'            adoRecSet.MoveNext
'        Next
'        Set adoRecSet = Nothing
'        Set connDB = Nothing
'        On Error GoTo 0
'    Else
        On Error Resume Next
        Call Setup_StrdNames
        Set M2Mrng = Range(Sheets("xxStdNames").Range("B1"), Sheets("xxStdNames").Range("B1").End(xlDown))
        Set M2Srng = Range(Sheets("xxStdNames").Range("D1"), Sheets("xxStdNames").Range("D1").End(xlDown))
        Set SYSrng = Range(Sheets("xxStdNames").Range("F1"), Sheets("xxStdNames").Range("F1").End(xlDown))
        For Each c In Range(Sheets("xxcalculations").Range("A1"), Sheets("xxcalculations").Range("A1").End(xlDown))
            c.Value = M2Mrng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, -1).Value
            c.Value = M2Srng.Find(what:=c.Value, lookat:=xlWhole).Offset(0, -1).Value
            c.Value = SYSrng.Find(what:=c.Value, lookat:=xlPart).Offset(0, -1).Value
        Next
        On Error GoTo 0
'    End If

    'assign spend to contracts/members
    '======================================================================================================================================
    If SetupSwitch = 2 Then
        Set MbrRng = Range(prsBKMRK.Offset(1, 0), prsBKMRK.Offset(MbrNMBR, 0))
        Set conrng = Range(ConTblBKMRK.Offset(1, 1), ConTblBKMRK.Offset(1, suppNMBR))
        MbrRng.Calculate
        prevCon = "x"
        For Each c In Range(Sheets("xxcalculations").Range("C1"), Sheets("xxcalculations").Range("C1").End(xlDown))
            If Not IsNumeric(c.Offset(0, -2)) Then
                If Not InStr(Trim(c.Value), prevCon) > 0 Then
                    If IsNumeric(Right(Trim(c.Value), 1)) Then
                        prevCon = Trim(c.Value)
                    ElseIf IsNumeric(Mid(Trim(c.Value), Len(Trim(c.Value)) - 1, 1)) Then
                        prevCon = Left(Trim(c.Value), Len(Trim(c.Value)) - 1)
                    Else
                        prevCon = Left(Trim(c.Value), Len(Trim(c.Value)) - 2)
                    End If
                    On Error GoTo errhndlNOCONTRACT
                    If AltPRS = True Then
                        For Pos = 1 To suppNMBR
                            If Trim(UCase(prevCon)) = Trim(UCase(AltPRSform.Controls("AltPRScon" & Pos).Text)) Then conpos = Pos
                        Next
                    Else
                        conpos = conrng.Find(what:=prevCon, lookat:=xlWhole).Column - 3
                    End If
                    MbrRng.Offset(0, conpos * 2).ClearContents
                End If
                On Error GoTo errhndlNOMBR
                mbraddr = MbrRng.Find(what:=c.Offset(0, -2).Value, lookat:=xlWhole, LookIn:=xlValues).Offset(0, conpos * 2).Address
                Sheets("Initiative Spend Overview").Range(mbraddr).Value = Sheets("Initiative Spend Overview").Range(mbraddr).Value + c.Offset(0, -1).Value
                On Error GoTo 0
            End If
NoMbr:  Next
        
NoCon:  Sheets("xxcalculations").Visible = False
        For i = 1 To suppNMBR
            MbrRng.Offset(0, i * 2).Replace what:=0, replacement:="", lookat:=xlWhole
        Next
        'Range(mbrrng.Offset(0, 2), mbrrng.Offset(0, suppNMBR * 2)).NumberFormat = "$0"
        Application.ScreenUpdating = True
        Range(prsBKMRK.Offset(MbrNMBR + 1, 1), prsBKMRK.Offset(MbrNMBR + 2, suppNMBR * 2 + 1)).Calculate
        'Range(mbrrng.Offset(2, 2), mbrrng.Offset(2, suppNMBR * 3)).Calculate
        If Not Trim(Sheets("notes").Range("AD2").Value) = "" Then
            Sheets("notes").Select
            Range("AD1:AF1").EntireColumn.Select
        Else
            Call FUN_TestForSheet("Initiative Spend Overview")
            prsBKMRK.Select
        End If
        
    Else
        For Each c In Range(Sheets("xxcalculations").Range("C2"), Sheets("xxcalculations").Range("C1").End(xlDown))
            If Not IsNumeric(c.Offset(0, -2)) Then
                If Trim(c.Value) <> prevCon Then
                    prevCon = Trim(c.Value)
                    conpos = conpos + 1
                    Sheets("PRS & UNSPSC").Range("C1").Offset(0, conpos * 2 - 1).Value = c.Value
                    mbrcnt = 0
                End If
                On Error GoTo errhndlNEWMBR
                mbraddr = Range(Sheets("PRS & UNSPSC").Range("B2").Offset(0, conpos * 2), Sheets("PRS & UNSPSC").Range("B1").Offset(0, conpos * 2).End(xlDown)).Find(what:=c.Offset(0, -2).Value, lookat:=xlWhole).Offset(0, 1).Address
                Sheets("PRS & UNSPSC").Range(mbraddr).Value = Sheets("PRS & UNSPSC").Range(mbraddr).Value + c.Offset(0, -1).Value
                On Error GoTo 0
            End If
NEWMBR: Next
        Sheets("xxcalculations").Visible = False
        Sheets("PRS & UNSPSC").Select
        For i = 1 To 10
            Range("C1").Offset(0, i * 2).Value = "=sum(" & Range(Range("C2").Offset(0, i * 2), Range("B1").Offset(0, i * 2 + 1).End(xlDown)).Address & ")"
            If Not IsEmpty(Range("B3").Offset(0, i * 2)) Then
                Range(Range("B2").Offset(0, i * 2), Range("B2").Offset(0, i * 2 + 1)).Columns.Hidden = False
                Range(Range("B2").Offset(0, i * 2), Range("B2").Offset(0, i * 2 + 1).End(xlDown)).Interior.Color = 14277081
                Range(Range("B2").Offset(0, i * 2), Range("B2").Offset(0, i * 2 + 1).End(xlDown)).Borders.LineStyle = xlContinuous
            Else
                Range(Range("B2").Offset(0, i * 2), Range("B2").Offset(0, i * 2 + 1)).Columns.Hidden = True
                Range(Range("B2").Offset(0, i * 2), Range("B2").Offset(0, i * 2 + 1)).Interior.Color = 14277081
                Range(Range("B2").Offset(0, i * 2), Range("B2").Offset(0, i * 2 + 1)).Borders.LineStyle = xlContinuous
            End If
        Next
        Range("Y1").Calculate
        If Not Range("A1:Y1").Interior.ColorIndex = 16 Then Call PRSUNSPSC_Template
        'Range("Z:AC").Columns.Hidden = True
        
    End If

On Error Resume Next
Unload AltPRSform
Application.StatusBar = False
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlReconnect:
On Error GoTo errhndlNORECSET
Set RDMconn = Nothing
Resume attempt2

errhndlNORECSET:
Resume exitPRS
exitPRS:
If SetupSwitch = 2 Then
    Sheets("notes").Range("AD:AD").ClearContents
    Sheets("notes").Range("AD1").Value = "UNASSIGNED PRS"
    Sheets("notes").Range("AD2").Value = "No PRS data found"
End If
Application.StatusBar = False
On Error Resume Next
Unload AltPRSform
Exit Sub

nowb:
Workbooks.Add
tempwb = 1
Resume Next

errhndlNOMBR:
mbrcnt = mbrcnt + 1
Range(Sheets("Notes").Range("AD1").Offset(mbrcnt, 0), Sheets("Notes").Range("AD1").Offset(mbrcnt, 2)).Value = Range(c.Offset(0, -2), c).Value
On Error GoTo 0
Resume NoMbr

errhndlNEWMBR:
mbrcnt = mbrcnt + 1
Range(Sheets("PRS & UNSPSC").Range("C1").Offset(mbrcnt, conpos * 2 - 1), Sheets("PRS & UNSPSC").Range("C1").Offset(mbrcnt, conpos * 2)).Value = Range(c.Offset(0, -2), c.Offset(0, -1)).Value
Resume NEWMBR

errhndlNOCONTRACT:
If Not MainCall = 1 Then MsgBox ("No associated contract found.  Please check your contract numbers in the tier info table and try again.")
Resume NoCon
'errhndlmbrsdone:
'Resume mbrsdone


End Sub
Sub Import_StdznIndex()

    Dim recset As New ADODB.Recordset
    Dim conn As New ADODB.Connection
    
    Set actvWB = ActiveWorkbook
    
    'import network sheet
    '===========================================================================================================================
    Call FUN_TestForSheet("SNA Standardization")
    
    Application.ScreenUpdating = False
'    connDB.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'         "Data Source=" & SNAengagePATH & ";" & _
'         "Extended Properties=""Excel 12.0;HDR=NO"";"
    conn.Open "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & Stdzn_Index_PATH & "\" & NetNm & ".xlsx"
    
    On Error GoTo errhndlNORECSET
    recset.Open "SELECT * FROM [Report$], conn, adOpenStatic, adLockReadOnly"
    recset.MoveFirst
    On Error GoTo 0

    actvWB.Activate
    Range("A1").CopyFromRecordset recset

    Set conn = Nothing
    Set recset = Nothing
    
    DoEvents
    For Each Wb In Workbooks
        If Wb.Name = "SNA Standardization Index.xlsx" Then Wb.Close (False)
    Next
    Application.ScreenUpdating = True

    If Not Range("B1:B2").Interior.Color = 15523812 Then Call SNAstandardization_Template
    Cells.WrapText = False
    
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNORECSET:
Range("A1").Value = "Unable to retrieve network tab"
Set conn = Nothing
QCerr = 1
Exit Sub



End Sub
Sub Import_CheatSheet()


Dim wdDoc         As Object
Dim TableNo       As Integer  'number of tables in Word doc
Dim iTable        As Integer  'table number index
    
    Set actvWB = ActiveWorkbook
    Call FUN_TestForSheet("CheatSheet")
    Cells.Clear
    CSfound = 0
    
    'loop through folders and find contract summary
    '---------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo errhndlNoPSC
    Set objFolder = objFSO.GetFolder(DATResourcePATH & "\PSC Cheat Sheets\" & ZeusForm.asscPSC.Value & "\") 'obviously replace
    On Error GoTo 0
    
    For Each ofile In objFolder.Files
        If InStr(ofile.Name, ".doc") > 0 Then

            'open word doc and extract table
            '---------------------------
            Set wdDoc = GetObject(ofile.Path)
            With wdDoc
                TableNo = wdDoc.Tables.Count
                If TableNo > 0 Then
                    wdDoc.Tables(1).Range.Copy
                    actvWB.Activate
                    Sheets("CheatSheet").Select
                    Cells.Clear
                    Range("A1").Select
                    ActiveSheet.Paste
                    wdDoc.Close (False)
                    Exit Sub
                Else
                    'MsgBox "Cheat sheet is not in standard format.  Requires manual extraction"
                    Set WordApp = CreateObject("word.Application")
                    WordApp.Documents.Open (ofile.Path)
                    WordApp.Visible = True
                    CSfound = 1
                End If
            End With
            
        ElseIf InStr(ofile.Name, ".xls") > 0 Then
            'MsgBox "Cheat sheet is not in standard format.  Requires manual extraction"
            ofilepath = ofile.Path
            Set NstdWB = Workbooks.Open(ofilepath)
            CSfound = 1
        End If
        
    Next ofile
    
    If Not CSfound = 1 Then
        Sheets("CheatSheet").Range("A1").Value = "No Cheat Sheet Available"
        Sheets("CheatSheet").Range("A1").Interior.ColorIndex = 3
        Sheets("CheatSheet").Range("A:A").Columns.AutoFit
        FolderName = DATResourcePATH & "\PSC Cheat Sheets\" & ZeusForm.asscPSC.Value
        retval = Shell("explorer.exe " & FolderName, vbNormalFocus)
    End If

Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoPSC:
Sheets("CheatSheet").Range("A1").Value = "No Cheat Sheet Available"
Sheets("CheatSheet").Range("A1").Interior.ColorIndex = 3
Sheets("CheatSheet").Range("A:A").Columns.AutoFit
FolderName = DATResourcePATH & "\PSC Cheat Sheets\"
retval = Shell("explorer.exe " & FolderName, vbNormalFocus)
Exit Sub


End Sub
Sub KeywordGenerator()

    'find range to be parsed
    '------------------------
    If Not ActiveSheet.Name = "Pricefile Keywords" Then Set keyrng = Selection
    
    Call FUN_TestForSheet("Pricefile Keywords")
    
    If Not IsEmpty(keyrng) Then
        Cells.Clear
        keyrng.Copy
        Range("A1").PasteSpecial Paste:=xlPasteValues
    End If
    
    Call HermesRunKeys
    
    'formatting
    '------------------------
    Range("A1").Select
    ActiveCell.EntireRow.Insert
    Range("A1").Value = "Item Descriptions"
    Range("B1").Value = "Single Word"
    Range("D1").Value = "Two Word"
    Range("F1").Value = "Three Word"
    Range("A1,B1,D1,F1").Interior.Color = 9420794
    Range("A1,B1,D1,F1").Borders.LineStyle = xlContinuous
    Range("A1,B1,D1,F1").Font.Bold = True
    Range("A:A").ColumnWidth = 125
    Range(Range("A2"), Range("A2").End(xlDown)).Interior.Color = 14281213
    Range(Range("C1"), Range("C2").End(xlDown)).Interior.Color = 4626167
    Range(Range("E1"), Range("E2").End(xlDown)).Interior.Color = 4626167
    Range(Range("G1"), Range("G2").End(xlDown)).Interior.Color = 4626167
    Range(Range("B1"), Range("B2").End(xlDown)).Borders.LineStyle = xlContinuous
    Range(Range("D1"), Range("D2").End(xlDown)).Borders.LineStyle = xlContinuous
    Range(Range("F1"), Range("F2").End(xlDown)).Borders.LineStyle = xlContinuous
    Range(Range("G1"), Range("G2").End(xlDown)).BorderAround ColorIndex:=1
    Range("A1:G1").Borders(xlEdgeBottom).Weight = xlThick


End Sub
Sub HermesRunKeys()

    HermesPhraseDensity 1, "B"
    HermesPhraseDensity 2, "D"
    HermesPhraseDensity 3, "F"
    
End Sub
Sub HermesPhraseDensity(nWds As Long, col As Variant)
    Dim astr()      As String
    Dim i           As Long
    Dim j           As Long
    Dim cell        As Range
    Dim sPair       As String
    Dim rOut        As Range

    With CreateObject("Scripting.Dictionary")
        .CompareMode = vbTextCompare
        For Each cell In Range("A1", Cells(Rows.Count, "A").End(xlUp))
            astr = Split(HermesLetters(cell.Value), " ")

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
        
        On Error Resume Next
        rOut.Columns(1).Value = Application.Transpose(.keys)
        rOut.Columns(2).Value = Application.Transpose(.items)

        rOut.Sort Key1:=rOut(1, 2), Order1:=xlDescending, _
                  Key2:=rOut(1, 1), Order1:=xlAscending, _
                  MatchCase:=False, Orientation:=xlTopToBottom, Header:=xlNo
        rOut.EntireColumn.AutoFit
        On Error GoTo 0
    End With
    
    
End Sub
Function HermesLetters(s As String) As String
    
    Dim i As Long

    For i = 1 To Len(s)
        Select Case Mid(s, i, 1)
            Case "A" To "ÿ", "a" To "ÿ", "A" To "Z", "a" To "z", "0" To "9", "'"
                HermesLetters = HermesLetters & Mid(s, i, 1)
            Case Else
                HermesLetters = HermesLetters & " "
        End Select
    Next i
    HermesLetters = WorksheetFunction.Trim(HermesLetters)
    
    
End Function

Sub Import_Dates()

Dim recset As New ADODB.Recordset
Dim conn As New ADODB.Connection

If SetupSwitch = 2 Then
    Set DatesSht = Sheets("index")
Else
    Set DatesSht = Sheets.Add
    Call Dates_Template
End If

If ZeusForm.extractSrch.Value = True Then

    '[TBD] delete dates table

Else

    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Stdzn_Index_PATH & "\" & NetNm & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    'conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Local_ETL_Data_PATH & "\" & "RDM" & "\" & NetNm & ".accdb"
    
    'SpendConn
    
    'SQLstr = "SELECT DISTINCT Name_Rolled_up_to_in_Report, IIF([Current_source] = 'RDM', RDM_Date_Range, Study_Date_Range), Current_Source FROM [Index_Tables$" & StdTbl & "] WHERE NOT Current_Source is null AND NOT Current_Source = 'Not Included' "
    sqlstr = "SELECT " & _
                "Name_Rolled_up_to_in_Report, " & _
                "IIF([Current_source] = 'RDM', min(left(RDM.Date_Range,instr(RDM.Date_Range,'-')-2)) + ' - ' + max(right(RDM.Date_Range,instr(RDM.Date_Range,'-')-2)), stdy.Date_Range), " & _
                "Current_Source " & _
             "FROM (([" & StdTbl & "$] as Std " & _
             "LEFT JOIN [" & RDMTbl & "$] as RDM ON RDM.Std_MID = Std.Std_MID) " & _
             "LEFT JOIN [" & StdyTbl & "$] as stdy ON Stdy.Std_MID = Std.Std_MID) " & _
             "WHERE NOT Current_Source is null AND NOT Current_Source = 'Not Included' AND IIF([Current_Source] = 'RDM', NOT RDM.Date_Range is null) " & _
             "GROUP BY Name_Rolled_up_to_in_Report, stdy.Date_Range, Current_Source"
            
'    SQLstr = "SELECT " & _
'                "Name_Rolled_up_to_in_Report, " & _
'                "min(format(left(RDM.Date_Range,instr(RDM.Date_Range,'-')-2), 'yyyymmdd')), " & _
'                "Current_Source " & _
'             "FROM (([" & StdTbl & "$] as Std " & _
'             "LEFT JOIN [" & RDMTbl & "$] as RDM ON RDM.Std_MID = Std.Std_MID) " & _
'             "LEFT JOIN [" & StdyTbl & "$] as stdy ON Stdy.Std_MID = Std.Std_MID) " & _
'             "WHERE NOT Current_Source is null AND NOT Current_Source = 'Not Included' " & _
'             "GROUP BY Name_Rolled_up_to_in_Report,RDM.Date_Range, stdy.Date_Range, Current_Source"
             
            '+ '-' + min(format(right(RDM.Date_Range,instr(RDM.Date_Range,'-')-2),'YYYY-MM-DD'))
            
    recset.Open sqlstr, conn, adOpenStatic, adLockReadOnly
    On Error Resume Next
    tmWB.Activate
    On Error GoTo 0
    Call FUN_TestForSheet("xxcalculations")
    Cells.Clear
    Range("A1").CopyFromRecordset recset
    Set conn = Nothing
    Set recset = Nothing
    
    'Get Primary date
    '------------------------------
    lstrw = FUN_lastrow("A")
    maxnmbr = 0
    If Application.CountIf(Range("C:C"), "RDM") >= 1 Then
        For i = 1 To lstrw
            If Range("C" & i).Value = "RDM" And Not Trim(Range("B" & i).Value) = "" Then
                currnmbr = Application.CountIf(Range("B:B"), Range("B" & i).Value)
                If currnmbr > maxnmbr Then
                    maxnmbr = currnmbr
                    prmyDate = Range("B" & i).Value
                End If
            End If
        Next
    Else
        For i = 1 To lstrw
            currnmbr = Application.CountIf(Range("B:B"), Range("B" & i).Value)
            If currnmbr > maxnmbr And Not Trim(Range("B" & i).Value) = "" Then
                maxnmbr = currnmbr
                prmyDate = Range("B" & i).Value
            End If
        Next
    End If
    
    DatesSht.Range("C24").Value = prmyDate
    
    'input names alternate ranges
    '------------------------------
    If Not Trim(DatesSht.Range("C26").Value) = "" Then Range(DatesSht.Range("C26"), DatesSht.Range("C22").End(xlDown).Offset(1, 2)).Clear
    For i = 1 To lstrw
        If Not Range("B" & i).Value = prmyDate And Not Trim(Range("B" & i).Value) = "" Then
            If Application.CountIf(Range(DatesSht.Range("C26"), DatesSht.Range("C22").End(xlDown)), Trim(Range("A" & i).Value)) > 0 Then
            
                Set preExistMbr = Range(DatesSht.Range("C26"), DatesSht.Range("C22").End(xlDown)).Find(what:=Trim(Range("A" & i).Value), lookat:=xlWhole)
                currdt = preExistMbr.Offset(0, 1).Value
                thisdt = Range("B" & i).Value
                
                If Format(Left(thisdt, InStr(thisdt, "-") - 2), "short date") < Format(Left(currdt, InStr(currdt, "-") - 2), "short date") Then
                    strtdt = Left(thisdt, InStr(thisdt, "-") - 2)
                Else
                    strtdt = Left(currdt, InStr(currdt, "-") - 2)
                End If
                
                If Format(Mid(thisdt, InStr(thisdt, "-") + 2, Len(thisdt)), "short date") > Format(Mid(currdt, InStr(currdt, "-") + 2, Len(currdt)), "short date") Then
                    enddt = Mid(thisdt, InStr(thisdt, "-") + 2, Len(thisdt))
                Else
                    enddt = Mid(currdt, InStr(currdt, "-") + 2, Len(currdt))
                End If
                preExistMbr.Offset(0, 1).Value = strtdt & " - " & enddt
            
            Else
                If DatesSht.Range("C22").End(xlDown).Offset(1, 0).Row > MbrBkmrk.Row + MbrNMBR Then DatesSht.Range("C22").End(xlDown).Offset(1, 0).EntireRow.Insert
                DatesSht.Range("C22").End(xlDown).Offset(1, 0).Value = Trim(Range("A" & i).Value)
                DatesSht.Range("C22").End(xlDown).Offset(0, 1) = Range("B" & i).Value    '& "  (" & Range("C" & i).Value & ")"
            End If
        End If
    Next
    If Not Trim(DatesSht.Range("C26").Value) = "" Then
        Set DatesRng = Range(DatesSht.Range("C26"), DatesSht.Range("C22").End(xlDown).Offset(0, 2))
        DatesRng.Borders.LineStyle = xlContinuous
        DatesRng.Borders.Color = 14277081
        DatesRng.HorizontalAlignment = xlLeft
        DatesRng.Font.ColorIndex = 1
        DatesRng.Font.Underline = False
        DatesRng.Font.Bold = False
        DatesRng.Font.Size = 8
    End If

End If


DatesSht.Select


End Sub

