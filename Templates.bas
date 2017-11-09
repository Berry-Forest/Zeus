Attribute VB_Name = "Templates"
Sub Dates_Template()


Range("C22").Value = "Date Ranges"
Range("C22:E22").Merge

Range("C23").Value = "Primary"
Range("C23:E23").Merge

Range("C24:E24").Merge

Range("C25").Value = "Alternate"
Range("C25:E25").Merge

Range("D26:E26").Merge

Range("C22:E25").HorizontalAlignment = xlCenter


End Sub
Sub Master_Template_Maintenance_SupplierFormulas()

'For each Supplier table (other than supplier 1) on activesheet
'*********************************

MftrNm_Lttr = "BG3"
InputCol_Lttr = "H"
DataCol_Lttr1 = "BP:BP"
'DataCol_Lttr2 = "BS:BS"
'DataCol_Lttr3 = "BT:BT"

For i = 1 To 9
    Range(InputCol_Lttr & "12").Offset(i * (150 + 8), 0).Formula = "=SUMIFS('Line Item Data'!" & Range(DataCol_Lttr1).Offset(0, i * 30).Address & ",'Line Item Data'!$P:$P,$B170,'Line Item Data'!$V:$V,'Line Item Data'!" & Range(MftrNm_Lttr).Offset(0, i * 30).Address & ",'Line Item Data'!$AI:$AI,""<>X"")"
    Range(InputCol_Lttr & "12").Offset(i * (150 + 8), 0).AutoFill Destination:=Range(Range(InputCol_Lttr & "12").Offset(i * (150 + 8), 0), Range(InputCol_Lttr & "12").Offset(i * (150 + 8) + 150 - 1, 0))
Next

End Sub
Sub Master_Template_Maintenance_GraphFormulas()

'For each Supplier table match rate graph (other than supplier 1) on activesheet
'*********************************

For i = 1 To 9
    mbrOffset = i * (150 + 8)
    Range("C9").Offset(mbrOffset, 0).Formula = "=(C" & 162 + mbrOffset & "-F" & 162 + mbrOffset & ")/C" & 162 + mbrOffset
    'Range(InputCol_Lttr & "12").Offset(i * (150 + 8), 0).Formula = "=SUMIFS('Line Item Data'!" & Range(DataCol_Lttr1).Offset(0, i * 30).Address & ",'Line Item Data'!$P:$P,$B170,'Line Item Data'!$V:$V,'Line Item Data'!" & Range(MftrNm_Lttr).Offset(0, i * 30).Address & ",'Line Item Data'!$AI:$AI,""<>X"")"
    'Range(InputCol_Lttr & "12").Offset(i * (150 + 8), 0).AutoFill Destination:=Range(Range(InputCol_Lttr & "12").Offset(i * (150 + 8), 0), Range(InputCol_Lttr & "12").Offset(i * (150 + 8) + 150 - 1, 0))
Next

End Sub
Sub ContractInfo_Template()

'    Dim conn As New ADODB.Connection
'    Dim RecSet As New ADODB.Recordset
    
    Call FUN_TestForSheet("Contract Info")
    
    'Format Cells
    '==========================================================================================================================================================
    
    'Headers
    '--------------------------------------------
    Range("A1:A2").Interior.ColorIndex = 1
    Range("A3").Value = "Portfolio Executive"
    Range("A4").Value = "Effective Date"
    Range("A5").Value = "Expiration Date"
    Range("A6").Value = "Standardization?"
    Range("A7").Value = "NovaPlus?"
    For i = 1 To 50
        Range("A" & 7 + i).Value = "Tier" & i
    Next
    
    For i = 1 To 10
        Range("A1").Offset(0, i).Value = "Contract " & i
    Next
    
    Range("M1").Value = "Network:"
    Range("M2").Value = "PSC:"
    Range("M3").Value = "Start Date:"
    Range("M4").Value = "End Date:"
    For i = 1 To 10
        Range("O1").Offset(i - 1, 0).Value = "Alt PRS " & i
    Next
    Range("Q1").Value = "Ad Hoc PSC"
    
    'Borders
    '--------------------------------------------
    Range("A1:K57,M1:N4,O1:P10,Q1:Q2").Borders.LineStyle = xlContinuous
    Range("A3:A57,M1:M4,N1:N4").Borders(xlEdgeRight).Weight = xlMedium
    Range("A7:K7,B2:K2,M2:N2,M4:N4").Borders(xlEdgeBottom).Weight = xlMedium
    
    'Colors
    '--------------------------------------------
    Range("A8:K57").Interior.Color = 65535
    Range("A3:K7").Interior.Color = 14281213
    Range("B2:K2").Interior.Color = 15849925
    Range("B1:K1").Interior.Color = 14857357
    Range("M1:M4").Interior.Color = 14336204
    Range("N1:N4").Interior.Color = 15523812
    Range("O1:O10,Q1").Interior.Color = 12566463
    Range("P1:P10,Q2").Interior.Color = 14277081
    Range("L1:L57").Interior.ColorIndex = 1
    
    Range("A:K").ColumnWidth = 13
    Range("M:M,O:O").ColumnWidth = 11
    Range("P:P").ColumnWidth = 14
    Range("L:L").ColumnWidth = 2
    Range("N:N,Q:Q").ColumnWidth = 30
    Range("A1:A57").Font.Size = 8
    Range("A3:A57,O1:O10,Q1,N:N").HorizontalAlignment = xlCenter

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'    'Pull psc list from edb
'    '==========================================================================================================================================================
'    SQLstr = "SELECT DISTINCT ATTRIBUTE_VALUE_NAME FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL WHERE ATTRIBUTE_NAME = 'PRODUCT SUB-CATEGORY' AND ATTRIBUTE_VALUE_STATUS = 'A' ORDER BY ATTRIBUTE_VALUE_NAME"
'
'    ConnStr = "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
'    conn.Open ConnStr
'
'    On Error GoTo errhndlNORECSET
'    RecSet.Open SQLstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
'    RecSet.MoveFirst
'    On Error GoTo 0
'
'    Range("R1").CopyFromRecordset RecSet
'    Range("R:R").Columns.Hidden = True
'
'    With Range("N2").Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=R1:" & Range("R1").End(xlDown).Address
'        .IgnoreBlank = False
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = ""
'        .InputMessage = ""
'        .ErrorMessage = "If you would like to use a non standard PSC you may choose Ad Hoc and enter the PSC in cell Q2 To the right."
'        .ShowInput = True
'        .ShowError = True
'    End With
'
'    Set RecSet = Nothing
'    Set conn = Nothing
'
'NoPSC:
'    'Pull network list from SNA engagement file
'    '==========================================================================================================================================================
'    For i = 1 To UBound(NtwkNmArray)
'        validationlist = validationlist & ", " & NtwkNmArray(i)
'    Next
'    validationlist = Right(validationlist, Len(validationlist) - 2)
'
'    With Range("N1").Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=validationlist
'        .IgnoreBlank = True
'        .InCellDropdown = True
'    End With
'
'
'Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'errhndlNORECSET:
'Range("A2").Value = "No PSC data found"
'Set conn = Nothing
'Resume NoPSC
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

End Sub
Sub PRSUNSPSC_Template()
    
    Call FUN_TestForSheet("PRS & UNSPSC")

    With Range("A1")
        .Value = "UNSPSC"
        .HorizontalAlignment = xlCenter
        .EntireColumn.ColumnWidth = 12
    End With
    Range("A1:B1").Merge
    Range("B1").EntireColumn.ColumnWidth = 40
    Range("A1:B1").BorderAround ColorIndex:=1, Weight:=xlMedium

    For i = 1 To 10
        Range("C1").Offset(0, i * 2).Value = "=sum(" & Range(Range("C2").Offset(0, i * 2), Range("B1").Offset(2, i * 2 + 1)).Address & ")"
        Range("C1").Offset(0, i * 2).NumberFormat = "$#,##0"
        Range("B1").Offset(0, i * 2).Value = "Contract " & i
        Range("B1:C1").Offset(0, i * 2).BorderAround ColorIndex:=1, Weight:=xlMedium
    Next
    Range("A1:Y1").Font.Bold = True
    Range("A1:Y1").Interior.ColorIndex = 16
    Range("D1:Y1").EntireColumn.ColumnWidth = 14
    Range("C1").Interior.ColorIndex = 1
    Range("C1").EntireColumn.ColumnWidth = 2
    Range("X1").Value = "Total PRS"
    Range("Y1").Formula = "=SUM(D1:W1)"
    Range("Y1").NumberFormat = "$#,##0"
    Range("X1:Y1").Interior.Color = 65535
    Range("X1:Y1").BorderAround ColorIndex:=1, Weight:=xlMedium


End Sub
Sub ScopeGuide_Template()

    Call FUN_TestForSheet("Scopeguide")

    Range("C1").Value = "In Scope"
    Range("I1").Value = "Out of Scope"
    Range("A1:K1").HorizontalAlignment = xlCenter
    Range("A2,G2").Value = "PIM Key"
    Range("B2,H2").Value = "Standard Manufacturer"
    Range("C2,I2").Value = "Standard Catalog Number"
    Range("D2,J2").Value = "PSC"
    Range("E2,K2").Value = "In or Out"
    Range("A1:K1").Interior.Color = 8421504
    Range("A1:K1").Font.Bold = True
    Range("A1:K1").BorderAround ColorIndex:=1 ', Weight:=xlThick
    Range("A2:K2").Interior.Color = 12632256
    Range("A2:K2").Borders.LineStyle = xlContinuous
    Range(Range("F1"), Range("E2").End(xlDown).Offset(0, 1)).Interior.ColorIndex = 1
    Range("K:K").ColumnWidth = 12


End Sub
Sub SNAstandardization_Template()

    Call FUN_TestForSheet("SNA Standardization")
    
    'Headers
    '------------------------------
    Range("A1").Value = "Network Run:"
    Range("A2").Value = "Date Range:"
    Range("A3").Value = "Hospital Name"
    Range("B3").Value = "Standardized Name"
    Range("C3").Value = "Study ID"
    Range("D3,G3").Value = "MID"
    Range("F3").Value = "Names To Be Included"
    Range("H3").Value = "Comments"
    
    'Formatting
    '------------------------------
    Range("A1:A2").Interior.Color = 14336204
    Range("B1:B2").Interior.Color = 15523812
    Range("A3:H3,E3:E500").Interior.Color = 15849925
    Range("I1:I503").Interior.ColorIndex = 1
    Range("A1:A2").Font.Bold = True
    Range("A3:H3").Font.Bold = True
    Range("A1:B2").Borders.LineStyle = xlContinuous
    Range("A3:H3").Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("A3:H3").Borders(xlEdgeTop).Weight = xlMedium
    Range("A3:H3").HorizontalAlignment = xlCenter
    Range("A:B,F:H").Columns.ColumnWidth = 30
    Range("C:D,G:G").Columns.ColumnWidth = 10
    Range("E:E,I:I").Columns.ColumnWidth = 2
    
    
End Sub
Sub Extract_Template()

Call FUN_TestForSheet("ASF Extract")

Range("A1").Value = "Original Sort"
Range("B1").Value = "CLINICAL Label"
Range("C1").Value = "PRODUCT SPEND CATEGORY"
Range("D1").Value = "NOV UNSPSC DESC"
Range("E1").Value = "PIM KEY"
Range("F1").Value = "HOSPITAL NAME"
Range("G1").Value = "HOSPITAL ITEM NUMBER"
Range("H1").Value = "HOSPITAL PRODUCT DESC"
Range("I1").Value = "HOSPITAL VENDOR NAME"
Range("J1").Value = "HOSPITAL VENDOR CATALOG NUMBER"
Range("K1").Value = "HOSPITAL MANUFACTURER NAME"
Range("L1").Value = "STANDARD MANUFACTURER NAME"
Range("M1").Value = "HOSPITAL MANUFACTURER CATALOG NUMBER"
Range("N1").Value = "STANDARD MANUFACTURER CATALOG #"
Range("O1").Value = "HOSPITAL QTY PER UOM"
Range("P1").Value = "HOSPITAL UOM"
Range("Q1").Value = "HOSPITAL PRICE"
Range("R1").Value = "Reported Current Usage"
Range("S1").Value = "Annual Usage Annualized"
Range("T1").Value = "TOTAL UOM UNITS"
Range("U1").Value = "Annualized Each Usage"
Range("V1").Value = "HOSPITAL UNIT PRICE"
Range("W1").Value = "CALCULATED HOSPITAL SPEND"
Range("X1").Value = "HOSPITAL TOTAL SPEND"
Range("Y1").Value = "MATCH CODE"
Range("Z1").Value = "NOV MANUFACTURER NAME"
Range("AA1").Value = "NOVATION CONTRACT ID"
Range("AB1").Value = "NOVATION CONTRACT"
Range("AC1").Value = "NOV PRODUCT DESC"
Range("AD1").Value = "NOV CATALOG PRODUCT"
Range("AE1").Value = "NOV UOM"
Range("AF1").Value = "NOV UNITS PER UOM"
Range("AG1").Value = "NOV UNIT PRICE"


End Sub
Sub DATxref_Template()


    'find DAT xref
    '============================================================================================================================================
    CursoryCheckFLG = 1
    Set xrefwb = Workbooks(FUN_OpenWBvar(ZeusPATH, "DATxref", FileName_PSC))
    CursoryCheckFLG = 0
    
    If wbfoundFLG = 0 Then
        Set xrefwb = Workbooks.Add
        ChDir ZeusPATH
        xrefwb.SaveAs ("CoreXref" & "(" & FileName_PSC & ")")
        Sheets.Add
        ActiveSheet.Name = "DATxref"
    Else
        On Error GoTo errhndlNOCORE
        Application.DisplayAlerts = False
        Sheets("DATxref").Copy After:=Sheets(Sheets.Count)
        Application.DisplayAlerts = True
        On Error GoTo 0
        Sheets("DATxref").Select
        Cells.Clear
    End If

NoCore:
    'Create Template
    '============================================================================================================================================
    Range("A2").Value = "Catalog Numbers"
    Range("B2").Value = "Item descriptions"
    Columns("A:B").Interior.ThemeColor = xlThemeColorAccent4
    Columns("A:B").Interior.TintAndShade = 0.799981688894314
    Range("C2").Value = "Catalog Numbers"
    Range("D2").Value = "Item descriptions"
    Columns("C:D").Interior.ThemeColor = xlThemeColorAccent3
    Columns("C:D").Interior.TintAndShade = 0.799981688894314
    Range("E2").Value = "Catalog Numbers"
    Range("F2").Value = "Item descriptions"
    Columns("E:F").Interior.ThemeColor = xlThemeColorAccent2
    Columns("E:F").Interior.TintAndShade = 0.799981688894314
    Range("G2").Value = "Catalog Numbers"
    Range("H2").Value = "Item descriptions"
    Columns("G:H").Interior.ThemeColor = xlThemeColorAccent5
    Columns("G:H").Interior.TintAndShade = 0.799981688894314
    Range("I2").Value = "Catalog Numbers"
    Range("J2").Value = "Item descriptions"
    Columns("I:J").Interior.ThemeColor = xlThemeColorAccent6
    Columns("I:J").Interior.TintAndShade = 0.799981688894314
    
    Range("A1").Value = "Supplier name"
    Range("A1:B1").Interior.ThemeColor = xlThemeColorAccent4
    Range("A1:B1").Interior.TintAndShade = 0.399975585192419
    Range("C1").Value = "Supplier name"
    Range("C1:D1").Interior.ThemeColor = xlThemeColorAccent3
    Range("C1:D1").Interior.TintAndShade = 0.399975585192419
    Range("E1").Value = "Supplier name"
    Range("E1:F1").Interior.ThemeColor = xlThemeColorAccent2
    Range("E1:F1").Interior.TintAndShade = 0.399975585192419
    Range("G1").Value = "Supplier name"
    Range("G1:H1").Interior.ThemeColor = xlThemeColorAccent5
    Range("G1:H1").Interior.TintAndShade = 0.399975585192419
    Range("I1").Value = "Supplier name"
    Range("I1:J1").Interior.ThemeColor = xlThemeColorAccent6
    Range("I1:J1").Interior.TintAndShade = 0.399975585192419
    Range("K1").Value = "Ect."
    Range("L1").Value = "..."
    
    Range("A2:L2").Columns.AutoFit


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::;
errhndlNOCORE:
Sheets.Add Before:=Sheets(1)
ActiveSheet.Name = "DATxref"
On Error GoTo 0
Resume NoCore
    
End Sub
Sub AdminFees_Template()
    
    Call FUN_TestForSheet("admin fees")

    Range("A1").Value = "Product ID"
    Range("B1").Value = "Contract UOM"
    Range("C1").Value = "Product Name"
    Range("D1").Value = "Base UOM"
    Range("E1").Value = "Base UOM Quantity"
    Range("F1").Value = "Price Program"
    Range("G1").Value = "Admin Fee Percent"
    Range("I1").Value = "Contract Number"
    Range("J1").Value = "Contract Name"
    Range("K1").Value = "Revenue Fee Detail"
    Range("A1:K1").Interior.Color = 15773696
    Range("H:H").Columns.ColumnWidth = 2
    Cells.WrapText = False
    Range("A1:K1").WrapText = True
    Range("A1:K1").Font.Bold = True
    Range("A1:K1").HorizontalAlignment = xlCenter
    Range("A1:K1").VerticalAlignment = xlCenter
    
End Sub
Sub Benchmark_Template()

Call FUN_TestForSheet("Best Market Price")

Range("A1").Value = "Part_Number"
Range("B1").Value = "PIM_Key"
Range("C1").Value = "10th% PLX_Best_Price"
Range("D1").Value = "25th% PLX_Best_Price"
Range("E1").Value = "50th% PLX_Best_Price"
Range("F1").Value = "Sample_Size"
Range("G1").Value = "NOV_Best_Price_Contract_ID"
Range("H1").Value = "Product_Desc"
Range("I1").Value = "Product_Spend_Category"
Range("J1").Value = "UOM"
Range("K1").Value = "UOM Qty"
Range("L1").Value = "10th% PLX_Best_EA_Price"
Range("M1").Value = "25th% PLX_Best_EA_Price"
Range("N1").Value = "50th% PLX_Best_EA_Price"
Range("O1").Value = "Benchmark Source"

Range("A1,B1,F1:K1,O1").Interior.Color = 12632256
Range("C1,L1").Interior.Color = 9420794
Range("D1,M1").Interior.Color = 14470546
Range("E1,N1").Interior.Color = 13082801

Range("A1:O1").Borders.LineStyle = xlContinuous
Range("A1:O1").HorizontalAlignment = xlCenter


End Sub
Sub QC_Template()

Call FUN_TestForSheet("QC")

'Headers
'===================================================================================================================================================

'Data
'-------------------------------
Range("A1").Value = "Location"
Range("B1").Value = "Help"
Range("C1").Value = "Description"
Range("E1").Value = "DAT Notes"
Range("F1").Value = "Format Review"
Range("G1").Value = "DAT Update"
Range("H1").Value = "PAT Review"

'color
'-------------------------------
Range("A1").Interior.Color = 15261367
Range("B1").Interior.Color = 65535
Range("C1").Interior.Color = 13082801
Range("D1").Interior.ColorIndex = 1
Range("E1").Interior.Color = 14994616
Range("F1").Interior.Color = 9420794
Range("G1").Interior.Color = 9420794
Range("H1").Interior.Color = 9737946

'format
'-------------------------------
Range("A1:H1").Borders.LineStyle = xlContinuous
Range("A1:H1").Font.Bold = True
Range("A1:H1").HorizontalAlignment = xlCenter


'title box
'===================================================================================================================================================

'Data
'-------------------------------
Range("J2").Value = "Network"
Range("J3").Value = "Initiative"
Range("J4").Value = "DAT (Task)"
Range("J5").Value = "DAT (QC)"
Range("J6").Value = "PAT (QC)"

Range("M2").Value = "Completed"
Range("M3").Value = "See Note"
Range("M4").Value = "Not Completed"
Range("M2").Interior.Color = 65280
Range("M3").Interior.Color = 65535
Range("M4").Interior.ColorIndex = 3

'format
'-------------------------------
Range("J2:J6").Font.Bold = True
Range("M2:M4").HorizontalAlignment = xlCenter
Range("M2:M4").BorderAround ColorIndex:=1
Range("J2:M6").BorderAround ColorIndex:=1, Weight:=xlMedium
Range("K2").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("K3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("K4").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("K5").Borders(xlEdgeBottom).LineStyle = xlContinuous

'Description Data
'===================================================================================================================================================

'Notes
'-------------------------------
Range("A4:A5").Value = "<TAB: Notes>"
Range("B4").Value = "Name of DAT completing the report must be noted on the notes tab."
Range("C4").Value = "DAT name included"

Range("B5").Value = "Variance % thresholds for Priceleveling, Supplier, and Benchmarking sections on the HCO tab that were used when standardizing UOM quantities must each be noted on the Notes tab. "
Range("C5").Value = "Variance ranges included"

'Spend Totals
'-------------------------------
Range("B8").Value = "Totals in all cells listed should match."

Range("A8").Value = "<TAB: Initiative Spend Overview TABLE: Market Share By Supplier>"
Range("C8").Value = "1) Market Share"

Range("A9").Value = "<TAB: Initiative Spend Overview TABLE: Current Purchases Benchmarking>"
Range("C9").Value = "2) Benchmark"

Range("A10").Value = "<TAB: Initiative Spend Overview TABLE: Supplier Reported Spend>"
Range("C10").Value = "3) PRS"

Range("A11").Value = "<TAB: Vizient Contracts - NC TABLE: Total Spend for each supplier table>"
Range("C11").Value = "4) Non Conversion"

Range("A12").Value = "<TAB: Vizient Contracts - Conv TABLE: Total Spend for each supplier table>"
Range("C12").Value = "5) Conversion"

Range("A13").Value = "<TAB: Line Item Data CELL: AJ3>"
Range("C13").Value = "6) Line Item Data"

'Index
'-------------------------------
Range("A16").Value = "<TAB: Index CELL: C8>"
Range("B16").Value = "The date should be present and the same as the date in the report's file name."
Range("C16").Value = "Report Create Date"

Range("A17").Value = "<TAB: Index CELL: C24 & C26+ >"
Range("B17").Value = "You can find this information on the network's tab in the ""SNA Standardization Index"" file. "
Range("C17").Value = "Member date ranges accurate"

Range("A18").Value = "<TAB: Index TABLE: Members>"
Range("B18").Value = "Required members present and member names standardized according to the network's tab in the file ""SNA Standarization Index""."
Range("C18").Value = "Required members present and standardized"

Range("A19").Value = "<TAB: Index TABLE: Keywords>"
Range("B19").Value = "Keywords used in data mining are captured in the keyword table."
Range("C19").Value = "Keywords present"

Range("A19").Value = "<TAB: Index TABLE: UNSPSC>"
Range("B19").Value = "Top ten codes and descriptions included.  If 10 aren't available then include as many as are."
Range("C19").Value = "UNSPSC data present"

Range("A20").Value = "<TAB: Impact Summary TABLE: Proposed Contract Terms and Conditions>"
Range("B20").Value = "Tier data and descriptions are accurate according to MPP.  AMC tiers are marked if not applicable to network.  Tier descriptions specific to other networks are removed.  Empty tier rows at the end are removed."
Range("C20").Value = "Tier data and descriptions are accurate"

Range("A21").Value = "<TAB: Impact Summary TABLE: Proposed Contract Terms and Conditions>"
Range("B21").Value = "Item count in column A of each supplier's pricing tab is correct."
Range("C21").Value = "Items on contract accurate"


'Initiative Spend Overview
'-------------------------------
Range("A24").Value = "<TAB: Initiative Spend Overview TABLE: Current Purchases Benchmarking>"
Range("B24").Value = "Member Current Spend totals are reasonable, as in larger HCO's have more spend than smaller HCO's. "
Range("C24").Value = "Member spend totals are reasonable"

Range("A25").Value = "<TAB: Initiative Spend Overview TABLE: Supplier Reported Spend (PRS)>"
Range("B25").Value = "For each supplier, total PRS is less than total spend"
Range("C25").Value = "PRS data is present and reasonable"

Range("A26").Value = "<TAB: Initiative Spend Overview TABLE: Supplier Reported Spend (PRS)>"
Range("B26").Value = "Suppliers are added to the table until the total spend for ""All Others"" is less than or equal to 5% (up to a maximum of 10 suppliers total)."
Range("C26").Value = """All Others"" total <= to 5%"


'Graphs
'-------------------------------
Range("A29:A30").Value = "<TAB: Initiative Spend Overview TABLE: Market Share Graph>"

Range("B29").Value = "Spend totals and percentages for each supplier on the graph match the spend totals and percentages in the two tables to the left."
Range("C29").Value = "Market share graph matches data"

Range("B30").Value = "Pie graph is clear, easily readable, and looks nice for when report is presented to members (fonts, sizing, position of supplier labels, position of graph, etc.)"
Range("C30").Value = "Market share graph is readable"

Range("A31:A32").Value = "<TAB: Initiative Spend Overview TABLE: Benchmarking Graph>"

Range("B31").Value = "Spend totals and percentages for each supplier on the graph match the spend totals and percentages in the two tables to the left."
Range("C31").Value = "Benchmarking graph matches data"

Range("B32").Value = "Pie graph is clear, easily readable, and looks nice for when report is presented to members (fonts, sizing, position of supplier labels, position of graph, etc.)"
Range("C32").Value = "Benchmarking graph is readable"


'Vizient Contracts - Conv
'-------------------------------
Range("A35").Value = "<TAB: Vizient Contracts - Conv TABLE: For each Supplier table>"
Range("B35").Value = "For each supplier the (Total Spend Needing Validation + Total New Proposed Spend + Total Spend Not Cross Referenced + Total Savings $) should equal Total Spend sum."
Range("C35").Value = "Proposed contract totals sum correctly"


'Line item data
'-------------------------------
Range("A38:A41").Value = "<TAB: Line Item Data  COLUMN: V >"

Range("B38").Value = "All contracted manufacturer names under ""Standard Manufacturer Names"" column are standardized to corresponding manufacturer name in supplier sections row 3, and manufacturer names in supplier sections are standardized to the supplier name as it appears in MPP."
Range("C38").Value = "Contracted Mftr names match MPP supplier name"

Range("B39").Value = "Manufacturer names standardized so that spend for each manufacturer is aggregated correctly within spreadsheet formulas. (Ex. Ethicon & Ethicon Inc are the same company and their name needs to match, else you will have spend for both Ethicon and Ethicon Inc instead of one or the other.)  Also check to make sure there are no inconsistencies in relation to patterns of associated catalog numbers.  (Ex. If you see 10 Medline catalog numbers with the prefix MYD- and another catalog number with the same prefix but it has different manufacturer name then the manufacturer name probably needs to be standardized to Medline.)"
Range("C39").Value = "Non contracted Mftr names standardized"

Range("B40").Value = "There cannot be more than one manufacturer name associated with the same catalog number.  (Ex. There cannot be an item with catalog number SNPA125P where the manufacturer name is Spiracur and another item with the same catalog number but with the manufacturer listed as Medline.)  A pivot table will help you to see manufacturer names per catalog number."
Range("C40").Value = "Only one Mftr name per catalog number"

Range("B41").Value = "There cannot be any blanks in the Standard Manufacturer Names column.  Also, there cannot be any manufacturer names left as ""Unknown"" unless you've done an exhaustive search to find the actual manufacturer of the item and still failed to find it.  Sometimes distributors are pulled in as manufacturers.  (Ex. Owens and Minor distributes products, they do not manufactur them.  If you see this or any other distributor in the Standard Manufacturer Names column you need to find the actual manufacturer of that item and change the name to that manufacturer.)"
Range("C41").Value = "No blank, unknown, or distributor Mftr names"

Range("A42:A43").Value = "<TAB: Line Item Data  COLUMN: X >"
Range("B42").Value = "Make sure all catalog numbers appear correctly standardized and normalized.  (Ex. If you see several items with the same catalog number and the catalog number doesn't have dashes, and then you see an item with the same catalog number but with a dash somewhere in it, then it probably needs to be standardized so that the dash is removed.)"
Range("C42").Value = "Catalog numbers standardized"

Range("B43").Value = "There cannot be any blanks in the Standard Manufacturer Catalog Numbers column.  Also, there cannot be any catalog numbers left as ""Unknown"" unless you've done an exhaustive search to find the actual catalog number for the item and still failed to find it."
Range("C43").Value = "No blank or unknown catalog numbers"

Range("A44").Value = "<TAB: Line Item Data  COLUMN: AO >"
Range("B44").Value = "If an item has a variance greater than the % threshold you are using for price leveling then the UOM quantity in column AB needs to be checked to make sure it's correct and adjusted if it's not.  If the UOM quantity is correct and the variance is still above your preset threshold then an x needs to be placed in column AI for that item.  If an item has an x in column AI but the variance exceeds +-200% then the item will still be flagged as needing to be resolved.  If the variance is greater than +-200% and you're certain you have the correct UOM quantity or you couldn't find the correct UOM quantity then you need to make a note for that item on the notes tab letting the PAT know that you did your research and couldn't come up with anything."
Range("C44").Value = "Price leveling variances have been resolved"

Range("A45").Value = "<TAB: Line Item Data  COLUMN: AV >"
Range("B45").Value = "If an item has a variance in the 10th percentile benchmark section greater than the benchmark threshold you are using, and you're sure the member UOM quantity is correct in column AB, then the UOM quantity in column K on the ""Best Market Price"" tab needs to be checked to make sure it's correct and adjusted if it's not.  If the UOM quantities are correct and the variance is still above your preset threshold then an x needs to be placed in column AI for that item.  If an item has an x in column AI but the variance exceeds +-200% then the item will still be flagged as needing to be resolved.  If the variance is greater than +-200% and you're certain you have the correct UOM quantity or you couldn't find the correct UOM quantity then you need to make a note for that item on the notes tab letting the PAT know that you did your research and couldn't come up with anything."
Range("C45").Value = "Tenth percentile variances resolved"

Range("A46").Value = "<TAB: Line Item Data  SUPPLIER SECTION: ""% Savings for Validation"" >"
str1 = "If an item has a variance in one of the supplier sections greater than the supplier threshold you are using, and your sure the member UOM quantity is correct in column AB, then the UOM quantity in column E of the supplier's pricing tab needs to be checked to make sure it's correct and adjusted if it's not.  If the variance is for an xref supplier, then before the UOM quantity on the supplier pricing tab is checked, the item being pulled in as a cross reference on the supplier's cross reference tab needs to be checked to make sure it's the best cross that can be used.  If the UOM quantities are correct and the variance is still above your preset threshold then an x needs to be placed in column AI for that item.  If an item has an x in column AI but the variance exceeds +-200% then the item will still be flagged as needing to be resolved."
str2 = " If the variance is greater than +-200% and you're certain you have the correct UOM quantity or you couldn't find the correct UOM quantity then you need to make a note for that item on the notes tab letting the PAT know that you did your research and couldn't come up with anything."
Range("B46").Value = str1 & str2
Range("C46").Value = "Supplier variances have been resolved"

Range("A47").Value = "<TAB: Line Item Data  SUPPLIER SECTION: ""10th % Variance"" >"
Range("B47").Value = "If an item has a variance in one of the benchmark supplier sections greater than the benchmark threshold you are using, and your sure the member UOM quantity is correct in column AB, then the UOM quantity in column K on the ""Best Market Price"" tab needs to be checked to make sure it's correct and adjusted if it's not.  If the UOM quantities are correct and the variance is still above your preset threshold then an x needs to be placed in column AI for that item.  If an item has an x in column AI but the variance exceeds +-200% then the item will still be flagged as needing to be resolved.  If the variance is greater than +-200% and you're certain you have the correct UOM quantity or you couldn't find the correct UOM quantity then you need to make a note for that item on the notes tab letting the PAT know that you did your research and couldn't come up with anything."
Range("C47").Value = "Supplier Benchmark variances resolved"

Range("A48").Value = "<TAB: Line Item Data  COLUMN: AC >"
Range("B48").Value = "Even if a UOM quantitiy isn't creating a significant variance, it's supposed to match up with the associated pkg description.  (Ex. If a pkg description in column AC is ""EA"" for each, then the UOM quantity should be 1.  Likewise if a pkg description is CA then the UOM quantity should probably be greater than 1) At the same time be aware that, although statistically less often, pkg descriptions can be wrong as well.  So if you've researched a UOM quantity and you strongly believe the UOM quantity for an item to be correct but the pkg description doesn't seem to match then that's fine, the point of the pkg description is to steer you in the right direction of the UOM quantitiy.  If that's the case though you might want to note it on the notes tab to CYA."
Range("C48").Value = "UOM pkg descriptions and UOM quantities match"

Range("A49").Value = "<TAB: Line Item Data>"
Range("B49").Value = "Items that are not within the scope of the PSC must be moved from the Line Item Data tab to the ""Items Removed"" tab.   (Ex. If your PSC is Patient Footwear and there's an item with the description ""epidural needle"" then that item is considered out of scope and needs to be removed.)  Also, items that are a match (""M"" in column F) to a contract that is not in your report or to an old contract being replaced by a contract in your report are considered out of scope."
Range("C49").Value = "Out of scope items removed"

Range("A50").Value = "<TAB: Line Item Data  SUPPLIER SECTION: ""Proposed Catalog #"" >"
Range("B50").Value = "If a contract is a novaplus contract then all items that cross in the supplier section for that contract on the Line Item Data tab that do not already show a Novaplus catalog number (usually denoted with a ""V"" prefix, but not always) must be checked to see if there is a Novaplus equivelent, and if there is then the Novaplus equivalent catalog number needs to be hardcoded as the proposed catalog number and highlighted in an orange color."
Range("C50").Value = "Convert to Novaplus codes where applicable"


'Pricing
'-------------------------------
Range("A53").Value = "<TAB: Supplier Pricing tab COLUMN: A>"
Range("B53").Value = "If there are duplicate catalog numbers then all but one need to be cut out and pasted to the removed items column out to the right.  You will need to consult with PAT to determine which price program they want to use."
Range("C53").Value = "Duplicate catalog numbers reconciled"

Range("A54").Value = "<TAB: Supplier Pricing tab COLUMN: J>"
Range("B54").Value = "$0 for a tier price will bring the tier used formula to $0 so any $0s must be changed to blanks."
Range("C54").Value = "Tier Used formulas correct"

Range("A55").Value = "<TAB: Supplier Pricing tab>"
Range("B55").Value = "If an item has no pricing in any tier greater than $0 then the catalog number must be cut out and moved to the Removed column out to the right."
Range("C55").Value = "$0 price items removed"

Range("A56").Value = "<TAB: Supplier Pricing tab>"
Range("B56").Value = "If there is pricing for a tier but the tier is not inlcuded in the contract per the tier info on the Index tab or the network is ineligible for that tier, then the tier column must be cut and pasted to the right and the best price formulas must be corrected so that they no longer pull in pricing form that tier."
Range("C56").Value = "Unqualified tiers removed"

Range("A57").Value = "<TAB: Supplier Pricing tab COLUMN: I>"
Range("B57").Value = "The date range of the items on each of the pricing tabs is not expired (Unless approved by PAT)."
Range("C57").Value = "Verified most current pricing being used"

'Cross Reference
'-------------------------------
Range("A60:A61").Value = "<TAB: Supplier xref tab COLUMN: A>"

Range("B60").Value = "Supplier Cross reference tabs are sorted correctly:  1st by Memer Catalog Number ascending, then by Source ascending, then by EA price ascending."
Range("C60").Value = "Data sorted correctly"

Range("B61").Value = "Make sure all catalog numbers appear correctly standardized and normalized.  (Ex. If you see several items with the same catalog number and the catalog number doesn't have dashes, and then you see an item with the same catalog number but with a dash somewhere in it, then it probably needs to be standardized so that the dash is removed.) Also make sure catalog numbers are converted to number so that they are recognized by other formulas within the report."
Range("C61").Value = "Xref codes cleansed"

'Admin fees
'-------------------------------
Range("A64").Value = "<TAB: Admin Fees>"
Range("B64").Value = "The fees in column G are correct in relation to the price program in column F as specified by the corresponding contract's revenue Fee detail in column K.  This takes into account not only net sales fees but also additional Novaplus/private label fees and fees specific to certain items or programs within that contract."
Range("C64").Value = "Correct Admin Fees are used"

'Best market Price
'-------------------------------
Range("A67").Value = "<TAB: Best Market Price>"
Range("B67").Value = "The data in the best market price tab should be sorted first by ""Part Number"" ascending, then by ""Sample Size"" descending."
Range("C67").Value = "Data correctly sorted"

'Overall
'-------------------------------
Range("A70:A72").Value = "<General Report>"

Range("B70").Value = "If automatic calulations turned off make sure report has been fully calculated before seding to QC or to PAT."
Range("C70").Value = "Report Fully Caluclated"

Range("B71").Value = "All data in all tabs should be set to Arial 8pt and zooms should be at 100%."
Range("C71").Value = "All fonts Arial 8pt and zooms at 100%"

Range("B72").Value = "Original extract, ASF extract, and Xref files used to create report are posted to the associated initiative folder under the corresponding network folder on the I: drive before sending to QC."
Range("C72").Value = "Original extract, ASF, and Xref posted"

Range("A73").Value = "<TAB: Line Item Data COLUMN: A>"
Range("B73").Value = "Member rows in the report are intact vs the raw data extract and no sort errors have occurred.  Any member name, PIM key, catalog number, ect. should correspond to the same original ID in column A as it does on the original extract."
Range("C73").Value = "Member data intact vs original extract"

Range("A74").Value = "<TAB: Initiative Spend Overview, Vizient Contracts - NC, Vizient Contracts - Conv, Line Item Data> "
Range("B74").Value = "This issue arises primarily in regard to the formulas in the tables on the Initiative Spend Overview tab and all the formulas to the right of the member data (column AM) on the Line Item Data tab.  Make sure all of the cells where formulas should be actually contain formulas and not a non-formulaic value."
Range("C74").Value = "No hardcoded formulas in dedicated formula cells"

'Formatting
'===================================================================================================================================================

'section headers
'-------------------------------
Range("C3").Value = "Notes"
Range("C7").Value = "Spend Totals Match On:"
Range("C15").Value = "Index"
Range("C23").Value = "Initiative Spend Overview"
Range("C28").Value = "Graphs"
Range("C34").Value = "Vizient Contracts - Conv"
Range("C37").Value = "Line Item Data"
Range("C52").Value = "Pricing"
Range("C59").Value = "Cross References"
Range("C63").Value = "Admin Fees"
Range("C66").Value = "Best Market Price"
Range("C69").Value = "Overall"

'Black line separator
'----------------------
Range("I:I").Interior.ColorIndex = 1

'bold section headers
'----------------------
Range("C3,C7,C15,C23,C28,C34,C37,C52,C59,C63,C66,C69").Font.Bold = True

Range("A4:A5, A8:A13, A16:A21, A24:A26, A29:A32, A35, A38:A50, A53:A57, A60:A61, A64, A67, A70:A74").Interior.Color = 15986394
Range("B4:B5, B8:B13, B16:B21, B24:B26, B29:B32, B35, B38:B50, B53:B57, B60:B61, B64, B67, B70:B74").Interior.Color = 7536379
Range("C4:C5, C8:C13, C16:C21, C24:C26, C29:C32, C35, C38:C50, C53:C57, C60:C61, C64, C67, C70:C74").Interior.Color = 15523812

Range("B4:H5, C8:H13, B16:H21, B24:H26, B29:H32,B35:H35,B38:H50,B53:H57,B60:H61, B64:H64, B67:H67, B70:H74").Borders.LineStyle = xlContinuous
Range("B8:B13").BorderAround ColorIndex:=1

Range("A:A").Columns.ColumnWidth = 85
Range("B:B").Columns.ColumnWidth = 132
Range("C:C").Columns.ColumnWidth = 45
Range("D:D").Columns.ColumnWidth = 2
Range("E:H").Columns.ColumnWidth = 18
Range("I:I").Columns.ColumnWidth = 2
Range("J:J").Columns.ColumnWidth = 10
Range("K:K").Columns.ColumnWidth = 35
Range("L:L").Columns.ColumnWidth = 2
Range("M:M").Columns.ColumnWidth = 13

Columns("A:B").Columns.Group
Range("A:B").Columns.Hidden = True

ActiveWindow.DisplayGridlines = False
Rows("2:2").Select
ActiveWindow.FreezePanes = True
Range("C1").Select

'protect cells
'----------------------
Cells.Locked = False
Range("C4:D5, C8:D13, C16:D21, C24:D26, C29:D32, C35:D35, C38:D50, C53:D57, C60:D61, C64:D64, C67:D67, C70:D74").Locked = True
'Range("D4:D5, D8:D13, D16:D21, D24:D26, D29:D32, D35:D35, D38:D50, D53:D57, D60:D61, D64:D64, D67:D67, D70:D74").Locked = True

ActiveSheet.Protect Password = "existentialism"


End Sub

Sub Master_Template()

Application.DisplayAlerts = False
Set mstrTplt = Workbooks.Open(TemplatePATH & "\" & MasterTemplate)
mstrTplt.SaveAs Filename:=ZeusPATH & "\" & "Blank_Template", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False


End Sub

Sub Master_Template_INPROGRESS()

'[INCOMPLETE]

Workbooks.Add

'add pricing and xref tabs
'======================================================================
Sheets(1).Name = "Pricing1"
Sheets(2).Name = "Pricing2"
Sheets(3).Name = "Pricing3"

For i = 4 To 10
    Sheets.Add
    ActiveSheet.Name = "Pricing" & i
Next
For i = 1 To 10
    Sheets.Add
    ActiveSheet.Name = "Cross Reference" & i
Next

'Add impact summary
'======================================================================
Sheets.Add
ActiveSheet.Name = "Impact Summary"

Range("A1").Value = Date
Range("A2").Value = "Product Spend Category"
Range("C1").Value = NetNm
Range("C2").Value = PSCVar




End Sub
Sub BRD_template()

'    On Error GoTo notOpen
'    Set wrdApp = GetObject("Word.Application")
''    If wrdApp Is Nothing Then
''        Set wrdApp = CreateObject("Word.Application")
''        Set wrdDoc = wrdApp.Documents.Open(TemplatePATH & "\" & BRDTemplate)
''        'CreateObject("Word.Application").Documents.Open (TemplatePATH & "\" & BRDTemplate)
''        'GetObject("Word.Application").Visible = True
''    Else
'     '   On Error GoTo notOpen
'        'Set wrdDoc = wrdApp.Documents(TemplatePATH & "\" & BRDTemplate)
'        'Set wrdDoc = wrdApp.Documents.Open(TemplatePATH & "\" & BRDTemplate)
'    'End If
'    'On Error Resume Next
'    wrdApp.Visible = True
'    wrdApp.Activate
'    Set wrdDoc = wrdApp.Documents.Open(TemplatePATH & "\" & BRDTemplate)
'
'
'Exit Sub
'':::::::::::::::::::::::::::::::::::::::::::::
'notOpen:
'CreateObject("Word.Application").Documents.Open (TemplatePATH & "\" & BRDTemplate)
On Error Resume Next
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = True
wrdApp.Activate
Set wrdDoc = wrdApp.Documents.Open(TemplatePATH & "\" & BRDTemplate)
Exit Sub


End Sub
Sub BRD_template_INPROGRESS()
'[INCOMPLETE]

Dim wrdApp As Word.Application
Dim wrdDoc As Word.Document
Dim SrcePath As String

'create new word doc
'-----------------------
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = True
    Set wrdDoc = wrdApp.Documents.Add

'add logo to header
'-----------------------
SrcePath = "C:\Users\Bforrest\Desktop\Zeus\1-Tools\Components\Icons\BRD Header.jpg"
wrdDoc.Sections.Item(1).headers(wdHeaderFooterPrimary).Range.InlineShapes.AddPicture (SrcePath)
With wrdDoc.Sections(1).headers(wdHeaderFooterPrimary).Range
    .Text = "Created by Analytics Services"
    .headers(wdHeaderFooterPrimary).Range.Text = Format(Date, "shortdate")
    .moveEnd wdCharacter, -1
    '.Font.Color =
With wrdDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "Footer goes here"


End With


    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

End Sub
Sub HCODetail_template()






End Sub
