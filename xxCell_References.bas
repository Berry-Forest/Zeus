Attribute VB_Name = "xxCell_References"
Public LID_Sheet As Worksheet
Public LIDHdrRow As Range
Public LIDdata_Rng As Range
Public data_sections() As String

Public UniqueID_BKMRK As Range
Public MbrCat_BKMRK As Range
Public MbrBench_BKMRK As Range
Public MbrUOM_Qty As Range
Public MbrUOM_Desc As Range
Public MbrUOM_Cost As Range

Public SuppCat_BKMRKs() As Range
Public SuppBench_BKMRKs() As Range
Public SuppUOM_Qtys() As Range
Public SuppUOM_Descs() As Range
Public SuppUOM_Costs() As Range

'Dim MemberCat_Header As String
'Dim SupplierCat_Header As String
'Dim BenchmarkData_Header As String
'Dim UOMqty_Header As String
'Dim UOMdesc_Header As String
'Dim UOMcost_Header As String
'Dim UniqueID_Header As String

Public SuppColNmbr As Integer
Public RefsCorrected As Boolean

Function FUN_SheetRefs() As Variant

    ReDim temp(1 To 2, 1 To 2)
    temp(1, 1) = "Name of sheet containing line item UOM data"          '<--not totally dynamic (search LIDRefs.Controls("lidrefs_txt1") )
    temp(2, 1) = "Row number of row containing UOM data headers"        '<--not totally dynamic (search LIDRefs.Controls("lidrefs_txt2") )
    temp(3, 1) = "Address for range containing all line item data."     '<--not totally dynamic (search LIDRefs.Controls("lidrefs_txt3") )

    temp(1, 2) = "Line Item Data"
    temp(2, 2) = 4
    temp(3, 2) = "$A$4:$MV$" & Range("X4").End(xlDown).Row
    
    FUN_SheetRefs = temp

End Function
Function FUN_HeaderRefs() As Variant

    ReDim temp(1 To 7, 1 To 2)
    temp(1, 1) = "Header name of column containing Unique ID for line items"
    temp(2, 1) = "Header name of column containing member catalog numbers"
    temp(3, 1) = "Header name of column containing 1st supplier catalog numbers"
    temp(4, 1) = "Header name of 1st column in benchmark sections"
    temp(5, 1) = "Header name of columns containing UOM quantities"
    temp(6, 1) = "Header name of columns containing UOM packaging descriptions"
    temp(7, 1) = "Header name of columns containing UOM package costs"
    
    temp(1, 2) = "Original Order"
    temp(2, 2) = "Standard Manufacturer Catalog #"
    temp(3, 2) = " - Proposed Catalog #"
    temp(4, 2) = "10th % Price UOM Cost"
    temp(5, 2) = "Quantity of Eaches per Unit of Measure"
    temp(6, 2) = "Unit of Measure Description"
    temp(7, 2) = "Unit of Measure Cost"
    
    
    FUN_HeaderRefs = temp
    
End Function
Sub Set_SectionRefs()

    ReDim data_sections(1 To 4)
    data_sections(1) = "Member"
    data_sections(2) = "Supplier"
    data_sections(3) = "Xref"
    data_sections(4) = "Benchmark"

End Sub
Sub set_data_refs()

'******************************************
'Use form data instead of array data cause array is default and user may have made ref changes or corrections
'******************************************

'Check to make sure data refs are correct
'================================================================================================================================================================================
    
    
    'Sheet ref
    '-----------------------------
    SheetChecks = FUN_SheetRefs
    ShtRefsNmbr = UBound(SheetChecks)
    
    On Error GoTo ERR_noLID
    Set LID_Sheet = Sheets(LIDRefs.Controls("lidrefs_txt1").Text)                       '<--[NOT DYNAMIC]
    On Error GoTo 0
    LID_Sheet.Columns.Hidden = False
    LID_Sheet.AutoFilterMode = False
    
    'Header Row
    '-----------------------------
    lidrow = LIDRefs.Controls("lidrefs_txt2").Text                                      '<--[NOT DYNAMIC]
    Set LIDHdrRow = LID_Sheet.Rows(lidrow & ":" & lidrow)
    
    'Total Data range
    '-----------------------------
    rng_val = LIDRefs.Controls("lidrefs_txt3").Text                                     '<--[NOT DYNAMIC]
    StartRng_Val = Cells(lidrow, Range(Left(rng_val, InStr(rng_val, ":") - 1)).Column).Address
    EndRng_Val = Mid(rng_val, InStr(rng_val, ":") + 1, Len(rng_val))
    Set LIDdata_Rng = Range(StartRng_Val & ":" & EndRng_Val)
    
    'header refs
    '-----------------------------
    HeaderChecks = FUN_HeaderRefs
    HdrRefsNmbr = UBound(HeaderChecks)
    
HeaderCheckStart:
    On Error GoTo ERR_noHeader
    For i = 1 To HdrRefsNmbr
        Set txtCTRL = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + i)
        Set LblCtrl = LIDRefs.Controls("lidrefs_lbl" & ShtRefsNmbr + i)
        'If Not Application.CountIf(LIDHdrRow, txtCTRL.Text) > 0 Then
        Set HeaderTest = LIDHdrRow.Find(what:=txtCTRL.Text, lookat:=xlPart)
    Next
    On Error GoTo 0
    
'Assign header names from LIDrefs form                                                  '<--[NOT DYNAMIC]
'================================================================================================================================================================================
    UniqueID_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 1).Text
    MemberCat_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 2).Text
    SupplierCat_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 3).Text
    BenchmarkData_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 4).Text
    UOMqty_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 5).Text
    UOMdesc_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 6).Text
    UOMcost_Header = LIDRefs.Controls("lidrefs_txt" & ShtRefsNmbr + 7).Text


'Define Section indexes
'================================================================================================================================================================================
    Call Set_SectionRefs

'Define Member data indexes
'================================================================================================================================================================================
    Set UniqueID_BKMRK = LIDHdrRow.Find(what:=UniqueID_Header, lookat:=xlPart)
    Set MbrCat_BKMRK = LIDHdrRow.Find(what:=MemberCat_Header, lookat:=xlPart)
    Set MbrBench_BKMRK = LIDHdrRow.Find(what:=BenchmarkData_Header, lookat:=xlPart)
    Set MbrUOM_Qty = LIDHdrRow.Find(what:=UOMqty_Header, lookat:=xlPart)
    Set MbrUOM_Desc = LIDHdrRow.Find(what:=UOMdesc_Header, lookat:=xlPart)
    Set MbrUOM_Cost = LIDHdrRow.Find(what:=UOMcost_Header, lookat:=xlPart)

'Define Supplier data indexes
'================================================================================================================================================================================
    If suppNMBR > 0 Then
        
        ReDim SuppCat_BKMRKs(1 To 1)
        ReDim SuppBench_BKMRKs(1 To 1)
        ReDim SuppUOM_Qtys(1 To 1)
        ReDim SuppUOM_Descs(1 To 1)
        ReDim SuppUOM_Costs(1 To 1)
        
        Set SuppStart = LIDHdrRow.Find(what:=SupplierCat_Header, lookat:=xlPart)
        SuppColNmbr = Range(SuppStart.Offset(0, 1), SuppStart.End(xlToRight)).Find(what:=SupplierCat_Header, lookat:=xlPart).Column - SuppStart.Column
        BenchOffset = Range(SuppStart, SuppStart.End(xlToRight)).Find(what:=BenchmarkData_Header, lookat:=xlPart).Column - SuppStart.Column
        UOMqtyOffset = Range(SuppStart, SuppStart.End(xlToRight)).Find(what:=UOMqty_Header, lookat:=xlPart).Column - SuppStart.Column
        UOMdescOffset = Range(SuppStart, SuppStart.End(xlToRight)).Find(what:=UOMdesc_Header, lookat:=xlPart).Column - SuppStart.Column
        UOMcostOffset = Range(SuppStart, SuppStart.End(xlToRight)).Find(what:=UOMcost_Header, lookat:=xlPart).Column - SuppStart.Column
        On Error GoTo 0
        
        For i = 1 To suppNMBR
            ReDim Preserve SuppCat_BKMRKs(1 To i)
            ReDim Preserve SuppBench_BKMRKs(1 To i)
            ReDim Preserve SuppUOM_Qtys(1 To i)
            ReDim Preserve SuppUOM_Descs(1 To i)
            ReDim Preserve SuppUOM_Costs(1 To i)
            Set SuppCat_BKMRKs(i) = SuppStart.Offset(0, (i - 1) * SuppColNmbr)
            Set SuppBench_BKMRKs(i) = SuppCat_BKMRKs(i).Offset(0, BenchOffset)
            Set SuppUOM_Qtys(i) = SuppCat_BKMRKs(i).Offset(0, UOMqtyOffset)
            Set SuppUOM_Descs(i) = SuppCat_BKMRKs(i).Offset(0, UOMdescOffset)
            Set SuppUOM_Costs(i) = SuppCat_BKMRKs(i).Offset(0, UOMcostOffset)
        Next
    End If
    
    
EndClean:
'================================================================================================================================================================================


    
Exit Sub
'::::::::::::::::::::::::::::::::
ERR_noLID:
    NewLIDName = Application.InputBox("Please input name of worksheet containing UOM data.", Type:=2)
    If NewLIDName = False Then
        Resume EndClean
    Else
        LIDRefs.Controls("lidrefs_txt1").Text = NewLIDName                              '<--[NOT DYNAMIC]
        Set LID_Sheet = Sheets(NewLIDName)
        Resume Next
    End If

ERR_noHeader:

    'if check fails more than 3 times then bring up form to define headers
    '-----------------------------
    CheckFails = CheckFails + 1
    If CheckFails > 2 Then
        If RefsCorrected = True Then
            RefsCorrected = False
            MsgBox ("Could not resolve user defined header references.  Please check your header references and try again.")
        Else
            RefsCorrected = True
            LIDRefs.Show (False)
        End If
        Resume EndClean
    End If
    
    'Check to make sure row 4 is header row
    '-----------------------------
    If Not RowCheck = True Then
        Row4_YesNo = MsgBox("Is the row containing headers row " & lidrow & "?", vbYesNo)
        If Not Row4_YesNo = vbYes Then
            NewHdrRow = Application.InputBox("Please input the row number of the row containing UOM data headers.  Headers must be on one row.", Type:=1)
            If Trim(NewHdrRow) = "" Then
                Resume EndClean
            Else
                LIDRefs.Controls("lidrefs_txt2").Text = NewHdrRow                       '<--[NOT DYNAMIC]
                LIDHdrRow = LID_Sheet.Rows(NewHdrRow & ":" & NewHdrRow)
                RowCheck = True
                Resume HeaderCheckStart
            End If
        End If
    End If
    
    'Input correct header name
    '-----------------------------
    NewHdrName = Application.InputBox("Please input " & LblCtrl.Caption & ".", Type:=2)
    If Trim(NewHdrName) = "" Then
        Resume EndClean
    Else
        txtCTRL.Text = NewHdrName
    End If

Resume


End Sub
