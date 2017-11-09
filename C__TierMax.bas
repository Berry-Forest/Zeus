Attribute VB_Name = "C__TierMax"

Sub Refresh_Suppliers()


Call AddSuppliers
Call Import_TierInfo
Call Import_Pricefile(, True)
Call Import_UNSPSC
Call Import_Benchmarking
Call Import_AdminFees
Call Calculate_Priceleveling
Call Import_PRS


End Sub
Sub StrtPlvl(control As IRibbonControl)

Call Calculate_Priceleveling

End Sub
Sub Calculate_Priceleveling()

On Error GoTo ERR_nonReport
Sheets("line item data").Visible = True
On Error GoTo 0
Call FUN_CalcOff  '>>>>>>>>>>
Sheets("line item data").Visible = True
ItmNmbr = Application.CountA(Range("X:X")) - 1
Range(Range("AM5"), Range("AM" & ItmNmbr + 4)).ClearContents
Range("AG:AJ").Calculate

For i = 1 To ItmNmbr + 1
'For i = 1 To 38198
    Application.StatusBar = "Calculating Priceleveling: " & i & " of " & ItmNmbr
    If Trim(Range("AM4").Offset(i, 0).Value) = "" Then
        MinEA = 1000000
        XoutMinEA = 1000000
        catnum = Range("X4").Offset(i, 0).Value
        mtchcats = Application.CountIf(Range("X:X"), catnum)
        
        'find min EA value
        '-----------------------------------
        On Error GoTo ERR_NxtCat
        Set srchstrt = Range("X4").Offset(i - 1, 0)
        For j = 1 To mtchcats
            Set srchstrt = Range(srchstrt, srchstrt.End(xlDown)).Find(what:=catnum, lookat:=xlWhole)
            currmin = srchstrt.Offset(0, 10).Value
            If currmin < MinEA Then
                XoutMinEA = currmin
                If Not LCase(srchstrt.Offset(0, 11).Value) = "x" Then MinEA = currmin
            End If
NxtCat:
        Next
        
        If MinEA = 1000000 Then MinEA = XoutMinEA
        If Left(MinEA, 1) = 0 And Len(MinEA) > 4 And Not MinEA = 0 Then
            MinEA = Format(MinEA, "$0.0000")
        Else
            MinEA = Format(MinEA, "$0.00")
        End If
        
        'populate min EA value
        '-----------------------------------
        On Error GoTo ERR_NxtCat2
        Set srchstrt = Range("X4").Offset(i - 1, 0)
        For j = 1 To mtchcats
            Set srchstrt = Range(srchstrt, srchstrt.End(xlDown)).Find(what:=catnum, lookat:=xlWhole)
            Range("AM" & srchstrt.Row).Value = MinEA
NxtCat2:
        Next
    End If
Next

Application.StatusBar = "Calcluating...Please Wait"
Range("AM:AQ").Calculate
If Not reportflg = 1 And Not RubiksFlg = 1 Then Call FUN_CalcBackOn  '>>>>>>>>>>
Application.StatusBar = False

Exit Sub
'::::::::::::::::::::::::
ERR_nonReport:
MsgBox "No Line Item Data tab found"
Exit Sub

ERR_NxtCat:
Resume NxtCat

ERR_NxtCat2:
Resume NxtCat2


End Sub
Sub Calculate_Priceleveling_Single(CalcRng As Range)


Call FUN_TestForSheet("line item data")
CalcRng.ClearContents
Range(CalcRng.Offset(0, -6), CalcRng.Offset(0, -3)).Calculate

ItmCnt = CalcRng.Count
MinEA = 1000000
For Each c In CalcRng.Offset(0, -5)
    If c.Value < MinEA Then
    XoutMinEA = c.Value
    If Not LCase(c.Offset(0, 1).Value) = "x" Then MinEA = c.Value
    End If
Next

If MinEA = 1000000 Then MinEA = XoutMinEA
'CalcRng.Value = Format(MinEA, "$0.0000")
If Left(MinEA, 1) = 0 And Len(MinEA) > 4 And Not MinEA = 0 Then
    CalcRng.Value = Format(MinEA, "$0.0000")
Else
    CalcRng.Value = Format(MinEA, "$0.00")
End If

Range(CalcRng.Offset(0, 1), CalcRng.Offset(0, 4)).Calculate


End Sub
Sub BRD()


Call BRD_template


End Sub
Sub METH_BRD_OLD()

    SetupSwitch = FUN_SetupSwitch(1)

    'create new word doc
    '-----------------------
    Application.ScreenUpdating = False
    Call BRD_template
    AppActivate "Microsoft Excel"
    
    'fill in top section
    '-----------------------
    wrdDoc.ContentControls("2695565873").Range.Text = NetNm
    wrdDoc.ContentControls("186335957").Range.Text = PSCVar
    If ZeusForm.NRSrch = True Then
        wrdDoc.ContentControls("2704741203").Range.Text = ZeusForm.asscStartDate.Text & "-" & ZeusForm.asscEndDate
    Else
        wrdDoc.ContentControls("2704741203").Range.Text = "One year from last submission per member"
    End If
    wrdDoc.Sections(1).headers(1).Range.Find.Execute findText:="7/1/2015", ReplaceWith:=Format(Date, "short date"), Forward:=True
    
    
    'extract and paste market share by supplier table
    '-----------------------
    Range(tmWB.Sheets("Current Market Share").Range("F4"), tmWB.Sheets("Current Market Share").Range("F4").Offset(MbrNMBR + 2, 0).End(xlToRight)).Copy
    wrdDoc.Range.GoTo(what:=wdGoToLine, which:=wdGoToAbsolute, Count:=12).PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    'wrdDoc.Range.Goto(What:=wdGoToLine, which:=wdGoToAbsolute, Count:=11).InsertParagraphAfter
    wrdDoc.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    wrdDoc.Tables(1).Rows.HeightRule = wdRowHeightAuto
    wrdDoc.Tables(1).Borders.OutsideLineStyle = wdLineStyleSingle
    'wrdDoc.Tables(1).Rows.Height = wrdDoc.Tables(1).Rows.Height + 1
 
    'convert to picture
    '-----------------------
'    wrdDoc.Tables(1).Range.CopyAsPicture
'    wrdDoc.Range.Goto(What:=wdGoToLine, which:=wdGoToAbsolute, Count:=11).PasteSpecial DataType:=wdPasteMetafilePicture
'    wrdDoc.Shapes(1).LockAspectRatio = True
'    wrdDoc.Shapes(1).Width = wrdDoc.PageSetup.TextColumns.Width
'
'    wrdDoc.Shapes(1).Left = wdShapeCenter
'
'    wrdDoc.Tables(1).Delete
'    For i = 1 To Int((MbrNMBR + 3) * 1.25)
'        wrdDoc.Range.Goto(What:=wdGoToLine, which:=wdGoToAbsolute, Count:=13).InsertParagraphAfter
'    Next
'
'    If wrdDoc.Shapes(1).Height > wrdDoc.Shapes(1).Width Then
'        wrdDoc.Shapes(1).Height = InchesToPoints(7.75)
'    Else
'        wrdDoc.Shapes(1).Width = InchesToPoints(7.5)
'    End If


    'Add rows to tables
    '-----------------------
    For i = 1 To suppNMBR - 1
        wrdDoc.Tables(2).Rows.Add
        wrdDoc.Tables(3).Rows.Add
    Next
    
    'fill in proposed supplier savings table
    '-----------------------
    For i = 1 To suppNMBR
        wrdDoc.Tables(2).cell(i + 2, 1).Range.Text = Sheets("impact summary").Range(Bkmrk).Offset(0, i).Value
        wrdDoc.Tables(2).cell(i + 2, 2).Range.Text = Trim(Replace(Sheets("Impact summary").Range(PropConBKMRK).Offset(-2, i * 4 + 2).Value, "Tier", ""))
        wrdDoc.Tables(2).cell(i + 2, 3).Range.Text = Format(Sheets("Impact summary").Range(PropConBKMRK).Offset(MbrNMBR + 1, i * 4 + 1).Value, "$#,##0.00")
        wrdDoc.Tables(2).cell(i + 2, 4).Range.Text = Format(Sheets("Impact summary").Range(PropConBKMRK).Offset(MbrNMBR + 1, i * 4 + 2).Value * 100, "%0.00")
    Next

    'fill in proposed supplier fees & rebates table
    '-----------------------
    For i = 1 To suppNMBR
        wrdDoc.Tables(3).cell(i + 2, 1).Range.Text = Sheets("impact summary").Range(Bkmrk).Offset(0, i).Value
        wrdDoc.Tables(3).cell(i + 2, 2).Range.Text = Sheets("Admin fees").Range("I:I").Find(what:=Sheets("impact Summary").Range(Bkmrk).Offset(1, i).Value).Offset(0, 2).Value
        wrdDoc.Tables(3).cell(i + 2, 6).Range.Text = Sheets("impact summary").Range(Bkmrk).Offset(6, i).Value
        wrdDoc.Tables(3).cell(i + 2, 7).Range.Text = Sheets("impact summary").Range(Bkmrk).Offset(5, i).Value
    Next
    
    'fill in negative impact member table
    '-----------------------
    For i = 1 To MbrNMBR
        For j = 1 To suppNMBR
            If Sheets("impact summary").Range(PropConBKMRK).Offset(i, j * 4 + 1).Value < 0 Then
                lastmbr = wrdDoc.Tables(5).Rows.Count
                If Not InStr(wrdDoc.Tables(5).cell(lastmbr, 1).Range.Text, Sheets("impact summary").Range(PropConBKMRK).Offset(i, 0).Value) > 0 Then
                    wrdDoc.Tables(5).Rows.Add
                    lastmbr = wrdDoc.Tables(5).Rows.Count
                    wrdDoc.Tables(5).cell(lastmbr, 1).Range.Text = Sheets("impact summary").Range(PropConBKMRK).Offset(i, 0).Value
                    wrdDoc.Tables(5).cell(lastmbr, 2).Range.Text = Format(Sheets("impact summary").Range(PropConBKMRK).Offset(i, 1).Value, "$#,###")
                End If
                wrdDoc.Tables(5).cell(lastmbr, j + 2).Range.Text = Format(Sheets("impact summary").Range(PropConBKMRK).Offset(i, j * 4 + 1).Value, "$#,###")
            End If
        Next
    Next
    wrdDoc.Tables(5).Rows(3).Delete
    
    'fill in negative impact product table
    '-----------------------
    Call FUN_TestForSheet("xxcalculations")
    Cells.ClearContents
    Range("A1").Value = "x"
    For i = 1 To suppNMBR
        Set negitmrng = Range(Sheets("Line Item Data").Range("AX3").Offset(0, (i - 1) * 18), Sheets("Line Item Data").Range("AX3").Offset(0, (i - 1) * 18).End(xlDown))
        For j = 1 To 5
            Sheets("xxcalculations").Range("A1").Offset((i - 1) * 5 + j, 0).Value = Application.WorksheetFunction.Small(negitmrng, j)
            Sheets("xxcalculations").Range("A1").Offset((i - 1) * 5 + j, 1).Value = negitmrng.Find(what:=Format(Application.WorksheetFunction.Small(negitmrng, j), "$#,##0.00"), LookIn:=xlValues).Address
        Next
    Next

    Set top5rng = Range(Sheets("xxcalculations").Range("A2"), Sheets("xxcalculations").Range("A1").End(xlDown))
    For i = 1 To 5
        nthaddress = top5rng.Find(what:=Application.WorksheetFunction.Small(top5rng, i), LookIn:=xlFormulas).Offset(0, 1).Value
        wrdDoc.Tables(6).cell(i + 2, 1).Range.Text = Sheets("Line Item Data").Range(nthaddress).Offset(0, -11).Value
        wrdDoc.Tables(6).cell(i + 2, 2).Range.Text = Sheets("Line Item Data").Range(nthaddress).Offset(0, -10).Value
        wrdDoc.Tables(6).cell(i + 2, 3).Range.Text = Sheets("Line Item Data").Range(nthaddress).Offset(0, -9).Value
        wrdDoc.Tables(6).cell(i + 2, 4).Range.Text = Format(Sheets("Line Item Data").Range(nthaddress).Offset(0, -1).Value / Sheets("Line Item Data").Range("X1").Value * 100, "%0.00")
        wrdDoc.Tables(6).cell(i + 2, 5).Range.Text = Format(Sheets("Line Item Data").Range(nthaddress).Value, "$#,##0.00")
    Next

    'fill in psotive benchmark product table
    '-----------------------
    Call FUN_TestForSheet("xxcalculations")
    Cells.ClearContents
    Range("A1").Value = "x"
    For i = 1 To suppNMBR
        Set positmrng = Range(Sheets("Line Item Data").Range("HW3").Offset(0, (i - 1) * 16), Sheets("Line Item Data").Range("HW3").Offset(0, (i - 1) * 16).End(xlDown))
        For j = 1 To 5
            posval = Format(Application.WorksheetFunction.Large(positmrng, j), "$#,##0.00")
            Sheets("xxcalculations").Range("A1").Offset((i - 1) * 5 + j, 0).Value = posval
            Sheets("xxcalculations").Range("A1").Offset((i - 1) * 5 + j, 1).Value = positmrng.Find(what:=posval, LookIn:=xlValues).Address
        Next
    Next

    Set top5rng = Range(Sheets("xxcalculations").Range("A2"), Sheets("xxcalculations").Range("A1").End(xlDown))
    For i = 1 To 5
        nthaddress = top5rng.Find(what:=Application.WorksheetFunction.Large(top5rng, i), LookIn:=xlFormulas).Offset(0, 1).Value
        nthpos = WorksheetFunction.RoundUp((top5rng.Find(what:=Application.WorksheetFunction.Large(top5rng, i), LookIn:=xlFormulas).Row - 1) / 5, 0)
        wrdDoc.Tables(7).cell(i + 2, 1).Range.Text = Sheets("Line Item Data").Range("HT1").Offset(0, (nthpos - 1) * 16).Value 'Sheets("Line Item Data").Range(Left(nthaddress, 3) & 1).Offset(0, -3).Value
        wrdDoc.Tables(7).cell(i + 2, 2).Range.Text = Sheets("Line Item Data").Range("AN" & Right(nthaddress, Len(nthaddress) - 4)).Offset(0, (nthpos - 1) * 18).Value
        wrdDoc.Tables(7).cell(i + 2, 3).Range.Text = Sheets("Line Item Data").Range("AO" & Right(nthaddress, Len(nthaddress) - 4)).Offset(0, (nthpos - 1) * 18).Value
        wrdDoc.Tables(7).cell(i + 2, 4).Range.Text = Format(Sheets("Line Item Data").Range(nthaddress).Offset(0, -1).Value / Sheets("Line Item Data").Range("X1").Value * 100, "%0.00")
        wrdDoc.Tables(7).cell(i + 2, 5).Range.Text = Format(Sheets("Line Item Data").Range(nthaddress).Value, "$#,##0.00")
    Next



'(from light blue benchmark)
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'    'fill in psotive benchmark product table
'    '-----------------------
'    Set top5rng = Range(Sheets("Line Item Data").Range("OA3"), Sheets("Line Item Data").Range("OA3").End(xlDown))
'    For i = 1 To 5
'        nthaddress = top5rng.Find(What:=Format(Application.WorksheetFunction.Large(top5rng, i), "$#,##0.00"), LookIn:=xlValues).Offset(0, 1).Address
'        wrdDoc.Tables(5).cell(i + 2, 1).Range.Text = Sheets("Line Item Data").Range(nthaddress).Offset(0, -11).Value
'        wrdDoc.Tables(5).cell(i + 2, 2).Range.Text = Sheets("Line Item Data").Range(nthaddress).Offset(0, -10).Value
'        wrdDoc.Tables(5).cell(i + 2, 3).Range.Text = Sheets("Line Item Data").Range(nthaddress).Offset(0, -9).Value
'        wrdDoc.Tables(5).cell(i + 2, 4).Range.Text = Format(Sheets("Line Item Data").Range(nthaddress).Offset(0, -1).Value / Sheets("Line Item Data").Range("X1").Value * 100, "%0.00")
'        wrdDoc.Tables(5).cell(i + 2, 5).Range.Text = Sheets("Line Item Data").Range(nthaddress).Value
'    Next
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
'Set otable = wrdDoc.bookmarks("propsuppfee")
    'wrdDoc.Tables(1).Select
    'wrdDoc.Range.InsertParagraphAfter '"Text Here"
  
'otable.Range.Select
'wrdDoc.Tables(otable).Rows.Add BeforeRow:=wrdDoc.Tables(otable).Rows(3)

'Selection.MoveDown Unit:=wdLine, Count:=1
'    wrdDoc.Tables(1).Range.collapse WdCollapseDirection.wdCollapseEnd
'
'    wrdDoc.Tables(1).Range.Select
'    'Selection.EndKey Unit:=wdColumn
'    wrdDoc.Tables(1).cell(Row:=suppNMBR + 2, Column:=1).Range.Select
'
'    'Selection.collapse
''    Selection.MoveDown Unit:=wdLine, Count:=1
'    wrdDoc.Tables(1).Range.GoTo what:=wdGoToTable, which:=wdGoToNext
'    'wrdDoc.Tables(1).MoveDown Unit:=wdLine, Count:=1
'
''    Selection.Range.TypeParagraph
'
'    For i = 1 To suppNMBR - 1
'        wrdDoc.Tables(2).Rows.Add BeforeRow:=wrdDoc.Tables(1).Rows(3)
'    Next
'    For i = 1 To suppNMBR - 1
'        wrdDoc.Tables(5).Rows.Add BeforeRow:=wrdDoc.Tables(1).Rows(3)
'    Next
'    For i = 1 To suppNMBR - 1
'        wrdDoc.Tables(6).Rows.Add BeforeRow:=wrdDoc.Tables(1).Rows(3)
'    Next

'Move to the end of the document
'oApp.Selection.Move Unit:=wdStory

'Selection.MoveDown Unit:=wdLine, Count:=8
'wrdDoc.Content.End.Select
'wrdDoc.Characters.Last.PasteSpecial DataType:=wdPasteMetafilePicture
'Set myTable = wrdDoc.Tables.Add(Range:=wrdDoc.Range(Start:=0, End:=0), NumRows:=1, NumColumns:=1)
'myTable.cell(1, 1).Range.InlineShapes.addpicture "C:\blah"
'wrdDoc.Tables(wrdDoc.Tables.Count).Range.InsertAfter " This is now the last sentence in paragraph one."
'ActiveDocument.Content.InsertAfter "end of document"
'wrdDoc.Range.Find.Execute(findText:="Title").Select
'wrdDoc.Tables(wrdDoc.Tables.Count).Range.Cut
'wrdDoc.Content.collapse Direction:=wdCollapseStart
'wrdDoc.Tables(wrdDoc.Tables.Count).Range.Paste

'Range(tmWB.Sheets("Current Market Share").Range("F4"), tmWB.Sheets("Current Market Share").Range("F4").Offset(MbrNMBR + 2, 0).End(xlToRight)).CopyPicture Appearance:=xlScreen, Format:=xlPicture
'wrdDoc.Characters.Last.Paste


Sheets("xxcalculations").Visible = False
Application.ScreenUpdating = True
Sheets("Impact summary").Select
wrdApp.Activate
wrdDoc.SaveAs ZeusPATH & NetNm & "_" & FileName_PSC & "_BRD_" & Replace(Format(Date, "short date"), "/", "-") & ".docx"
Set wrdApp = Nothing
Set wrdDoc = Nothing


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::




End Sub

Sub METH_NovaPlus()

'Find Novaplus codes if a mfg is Novaplus
'*****************************************************

    If Application.CountIf(Range(ConTblBKMRK.Offset(6, 1), ConTblBKMRK.Offset(6, suppNMBR)), "ü") > 0 Then
        Set NovaWB = Workbooks(FUN_OpenWBvar(ZeusPATH & "\1-DB shortcuts\", "NOVAPLUS", ".xlsx"))
        tmWB.Activate
        Sheets("line item data").Select
        
        For i = 0 To suppNMBR - 1
            If ConTblBKMRK.Offset(6, i + 1).Value = "ü" Then
                For Each c In Range(Range("BG5").Offset(0, i * 30), Range("BG5").Offset(0, i * 30).End(xlDown))
                    If Not c.Value = "-" Then
                        If Not InStr(UCase(c.Value), "V") = 1 Then
                            findnmbr = c.Offset(0, Range("X1").Column - c.Column).Value
                            On Error GoTo ERR_NxtItm
                            novanmbr = NovaWB.Sheets("FEQ Template").Range("E:E").Find(what:=findnmbr, lookat:=xlWhole).Offset(0, -3).Value
                            On Error GoTo 0
                            c.Value = novanmbr
                        Else
                            c.Value = c.Value 'to hardcode it if it has a V
                        End If
                    End If

NxtItm:         Next
            End If
        Next
        
        NovaWB.Close (False)
    End If

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_NxtItm:
On Error GoTo 0
Resume NxtItm



End Sub
Sub StandardizeMfg()

'does not require sorting
'***************************************************************************


If MainCall = 1 And Not QCFlg = 1 Then
    spendsheet = "Spend Search"
    HdrRw = 1
Else
    spendsheet = "Line Item Data"
    HdrRw = 4
    For Each c In Sheets(spendsheet).Range("V5:V" & Range("A4").End(xlDown).Row)
        If c.Interior.Color = 16711935 Then c.Interior.ColorIndex = 0
    Next
End If

Sheets(spendsheet).Select
ActiveSheet.AutoFilterMode = False
LastRow = FUN_lastrow("A")
Set MftrRng = Sheets(spendsheet).Range("V5:V" & LastRow)
sData = spendsheet & "!R" & HdrRw & "C22:R" & LastRow & "C24"
sNbr = Sheets(spendsheet).Range("X" & HdrRw).Value
sName = Sheets(spendsheet).Range("V" & HdrRw).Value
QCChkFlg = True

EndPivotCreate:

'Setup for loop
'==============================================================================================
    On Error GoTo errhndlMFGPIVOT
    Sheets("supplier pivot").Visible = True
    Sheets("supplier pivot").Select
    On Error GoTo 0

'Refresh Pivot (won't change columns if column changes have already been made)
'=======================================================================================
If Not pivotcreateFLG = 1 Then

    'Make sure table is pointing at the right data and is refreshed
    '----------------------------------
    Application.DisplayAlerts = False
    Sheets("supplier pivot").Select
    Range("A1").Select

    On Error Resume Next
    ActiveSheet.PivotTables(1).ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sData, Version:=xlPivotTableVersion14)
    ActiveSheet.PivotTables(1).PivotCache.Refresh
    On Error GoTo 0
        
    'Hide all current fields
    '----------------------------------
'    For Each f In ActiveSheet.PivotTables(1).PivotFields
'        If Not (f.Name = sName Or f.Name = sNbr Or f.Orientation = xlHidden) Then
'            On Error Resume Next
'            f.Orientation = xlHidden
'        End If
'    Next
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    With ActiveSheet.PivotTables(1).PivotFields(sNbr)
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables(1).PivotFields(sName)
        .Orientation = xlRowField
        .Position = 2
    End With

End If

'<LOOP>
'==============================================================================================
Dim StMfg As String
Dim MfgNmbr As Integer
Range("A1").End(xlDown).Select

If Not IsEmpty(ActiveCell.Offset(1, 1)) Then
    Range("B1").End(xlDown).Offset(1, -1).Select
    lastmftr = ActiveCell.End(xlUp).Value
    Range("A1").End(xlDown).Select

    Do
        DoEvents
        Application.StatusBar = "Checking manufacturer pivot: (Row) " & ActiveCell.Row
        
        'Count number of mfgs with the same catnum
        '-------------------------------
        If Not ActiveCell.Value = lastmftr Then
            Range(Selection, Selection.End(xlDown).Offset(-1, 0)).Select
        Else
            Range(ActiveCell, Range("B1").End(xlDown).Offset(0, -1)).Select
        End If
        MfgNmbr = Selection.Count
            
        'Determine greatest common assoc supplier
        '-------------------------------
        CatNumBkmrk = ActiveCell.Address
        Range("C:C").ClearContents
        Range("C1").Formula = "=CountIf(" & "'" & spendsheet & "'" & "!V:V," & ActiveCell.Offset(0, 1).Address(0, 0) & ")"
        Range("C1").AutoFill Destination:=Range("C1:C" & MfgNmbr)
        Range("D1").Formula = "=countif(C1,Max(C:C))"
        Range("D1").AutoFill Destination:=Range("D1:D" & MfgNmbr)
        Range("C:D").Calculate
        mfgMtchCats = Application.sum(Range("C:C"))
        Range("D:D").Select
        Selection.Find(what:="1", After:=ActiveCell, LookIn:=xlValues, _
            lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Select
        MfgRef = ActiveCell.Row
        Range(CatNumBkmrk).Select
        FndCatNum = ActiveCell.Value
        StMfg = ActiveCell.Offset(MfgRef - 1, 1)
        
        'Change mfg on HCO detail tab
        '-------------------------------
        Sheets(spendsheet).Select
        Range("X2").Select
        mtchcats = Application.CountIf(Range("X:X"), FndCatNum)
        FndCatNumCNT = 0
        Do
        FndCatNumCNT = FndCatNumCNT + 1
            
            Range(ActiveCell, Range("X" & LastRow)).Find(what:=FndCatNum, After:=ActiveCell, LookIn:=xlFormulas, lookat:=xlWhole, SearchDirection:=xlNext, MatchCase:=False).Select
            If Not ActiveCell.Offset(0, -2).Value = StMfg Then
                If MainCall = 1 Or IndivFlg = 1 Then
                    ActiveCell.Offset(0, -2).Value = StMfg
                    ActiveCell.Offset(0, -2).Interior.Color = 65535
                ElseIf QCFlg = True Then
                    QCChkFlg = False
                    ActiveCell.Offset(0, -2).Interior.Color = 16711935
                End If
            End If
            
        Loop Until FndCatNumCNT = mtchcats
    
        'Go to next supplier
        '-------------------------------
        Sheets("supplier pivot").Select
        If Not ActiveCell.Value = lastmftr Then
            Selection.End(xlDown).Select
            If Not IsEmpty(ActiveCell.Offset(1, 0)) Then
                Selection.End(xlDown).Select
            End If
        Else
            Exit Do
        End If
    
    Loop Until IsEmpty(ActiveCell.Offset(1, 1))

End If

On Error Resume Next
ActiveSheet.PivotTables(1).PivotCache.Refresh
On Error GoTo 0
Application.StatusBar = False
    

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlMFGPIVOT:
Resume createPivot:
createPivot:
'Create initial Pivot for catalog and name standardization
'===========================================================================================
    Application.StatusBar = "Creating Pivot table...Please wait"
    pivNm = "SupplierPivotTable"

    On Error GoTo errhndlAltPivotName
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        sData, Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="", TableName:=pivNm, DefaultVersion:=xlPivotTableVersion14
    On Error GoTo 0
    'Add mfg catalogs
    With ActiveSheet.PivotTables(pivNm).PivotFields(sNbr)
        .Orientation = xlRowField
        .Position = 1
    End With
    'Add mfg names
    With ActiveSheet.PivotTables(pivNm).PivotFields(sName)
        .Orientation = xlRowField
        .Position = 2
    End With
    'Format table
        'Tabular
    ActiveSheet.PivotTables(pivNm).RowAxisLayout xlTabularRow
    With ActiveSheet.PivotTables(pivNm)
        .ColumnGrand = False
        .RowGrand = False
    End With
        'No +/-
    ActiveSheet.PivotTables(pivNm).ShowDrillIndicators = False
        'No subtotals for either field
    ActiveSheet.PivotTables(pivNm).PivotFields( _
        sName).Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    ActiveSheet.PivotTables(pivNm).PivotFields( _
        sNbr).Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
        'Rename Sheet
    ActiveSheet.Name = "Supplier Pivot"
    Columns("A:A").Select
    Selection.ColumnWidth = 27.43
    Columns("B:B").Select
    Selection.ColumnWidth = 27.43
    pivotcreateFLG = 1
'Application.ScreenUpdating = True
Application.StatusBar = False
GoTo EndPivotCreate

errhndlAltPivotName:
    pivCnt = pivCnt + 1
    pivNm = "SupplierPivotTableAlt" & pivCnt
    If pivCnt > 3 Then
        QCChkFlg = False
        Exit Sub
    End If
Resume

''Create initial Pivot for catalog and name standardization
''===========================================================================================
'    Application.StatusBar = "Creating Pivot table...Please wait"
'    If MainCall = 1 Then
'        sData = spendsheet & "!R1C12:R" & lastrow & "C14"
'        sNbr = Range("N1").Value
'        sName = Range("L1").Value
'    Else
'        sData = spendsheet & "!R2C12:R" & lastrow & "C14"
'        sNbr = "Standard Manufacturer Catalog #"
'        sName = "Standard Manufacturer Name"
'    End If
'    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        sData, Version:=xlPivotTableVersion14). _
'        CreatePivotTable TableDestination:="", TableName:="SupplierPivotTableAlt", DefaultVersion:=xlPivotTableVersion14
'    'Add mfg catalogs
'    With ActiveSheet.PivotTables("supplierPivotTableAlt").PivotFields(sNbr)
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    'Add mfg names
'    With ActiveSheet.PivotTables("supplierPivotTableAlt").PivotFields(sName)
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'    'Format table
'        'Tabular
'    ActiveSheet.PivotTables("SupplierPivotTableAlt").RowAxisLayout xlTabularRow
'    With ActiveSheet.PivotTables("SupplierPivotTableAlt")
'        .ColumnGrand = False
'        .RowGrand = False
'    End With
'        'No +/-
'    ActiveSheet.PivotTables("SupplierPivotTableAlt").ShowDrillIndicators = False
'        'No subtotals for either field
'    ActiveSheet.PivotTables("SupplierPivotTableAlt").PivotFields( _
'        sName).Subtotals = Array(False, False, False, False, False, _
'        False, False, False, False, False, False, False)
'    ActiveSheet.PivotTables("SupplierPivotTableAlt").PivotFields( _
'        sNbr).Subtotals = Array(False, False, False, False, False, _
'        False, False, False, False, False, False, False)
'        'Rename Sheet
'    ActiveSheet.Name = "Supplier Pivot Alt"
'    Columns("A:A").Select
'    Selection.ColumnWidth = 27.43
'    Columns("B:B").Select
'    Selection.ColumnWidth = 27.43
'    pivotcreateFLG = 1
''Application.ScreenUpdating = True
'Application.StatusBar = False
'SupplierPivotAltFLG = 1
'Resume EndPivotCreate



End Sub
Sub StdzContractedMftrs()

LastRow = Sheets("Line item data").Range("A4").End(xlDown).Row
Set MftrRng = Sheets("Line item data").Range("V5:V" & LastRow)
For i = 1 To suppNMBR
    SuppNm = Sheets("Line item data").Range("BG3").Offset(0, (i - 1) * 30)
    Set SuppRng = Sheets("Line item data").Range("BG5:BG" & LastRow).Offset(0, (i - 1) * 30)
    SuppRng.Calculate
    For Each c In SuppRng
        If Not c.Value = "-" Then
            If Range("X" & c.Row).Value = c.Value Then
                If Range("F" & c.Row) = "M" Then
                    If Not Range("V" & c.Row).Value = SuppNm Then
                        MftrRng.Replace what:=Range("V" & c.Row).Value, replacement:=SuppNm, lookat:=xlWhole, MatchCase:=False
                        'Exit For
                    End If
                End If
            End If
        End If
    Next
Next



End Sub
Sub AddSuppliers()


    'Delete supplier summary columns and tables
    '==================================================================================
    If Not prsBKMRK.End(xlToRight).Column = prsBKMRK.Offset(0, 3).Column Then
        Range(MSGraphBKMRK.Offset(0, 3), MSGraphBKMRK.End(xlToRight).Offset(MbrNMBR + 1, -3)).Delete Shift:=xlToLeft
        Range(prsBKMRK.Offset(0, 3), prsBKMRK.End(xlToRight).Offset(MbrNMBR + 2, -1)).Delete Shift:=xlToLeft
    End If
    
    'Delete tables on NC Tab
    '---------------------------
    Call FUN_TestForSheet("Vizient Contracts - NC")
    Rows.Hidden = False
    Range(NonConBKMRK.Offset(MbrNMBR + 5, -1), NonConBKMRK.Offset(Range("B" & Rows.Count).End(xlUp).Row + 2, 0)).EntireRow.Delete Shift:=xlToLeft
    On Error Resume Next
    For i = 1 To ActiveSheet.Shapes.Count
        If ActiveSheet.Shapes(i).Type = 3 And Not ActiveSheet.Shapes(i).Name = "MatchGraph1" Then ActiveSheet.Shapes(i).Delete
    Next
    On Error GoTo 0
    
    'Delete tables on Conv Tab
    '---------------------------
    Call FUN_TestForSheet("Vizient Contracts - Conv")
    Rows.Hidden = False
    Range(ConvBKMRK.Offset(MbrNMBR + 5, -1), ConvBKMRK.Offset(Range("B" & Rows.Count).End(xlUp).Row + 2, 0)).EntireRow.Delete Shift:=xlToLeft
    On Error Resume Next
    For Each shp In ActiveSheet.Shapes
        If shp.Type = 8 And Not shp.Name = "ConvChkBx1" Then
            shp.Delete
        ElseIf shp.Type = 3 And Not shp.Name = "MatchGraph1" Then
            shp.Delete
        End If
    Next
    On Error GoTo 0
    
    'Add Contracted Suppliers
    '==================================================================================
    For i = 2 To suppNMBR
        
        'copy formulas in LID supplier section if none found
        '----------------------------------------
        Set suppstrt = LIDSuppBKMRK.Offset(1, (i - 1) * 30)
        If IsEmpty(suppstrt) Then
            Range(LIDSuppBKMRK.Offset(1, 0), LIDSuppBKMRK.Offset(1, 29)).Copy
            suppstrt.PasteSpecial xlPasteAll
            Range(suppstrt, suppstrt.Offset(0, 29)).Replace what:=FUN_SuppName(1), replacement:=suppstrt.Offset(-2, 0).Value, lookat:=xlPart
        End If
        
        If IsEmpty(suppstrt.Offset(1, 0)) Then
            Range(suppstrt, suppstrt.Offset(0, 29)).AutoFill Destination:=Range(suppstrt, suppstrt.Offset(ItmNmbr - 1, 29))
        End If

        'Add suppliers to summary tabs
        '----------------------------------------
        Call AddSupplier_Col(False)
        
        Sheets("Vizient Contracts - NC").Select
        Call AddSupplier_Summary(i, NonConBKMRK)
        
        Sheets("Vizient Contracts - Conv").Select
        Call AddSupplier_Summary(i, ConvBKMRK)
        
        'Copy Conversion check box
        '----------------------------------------
        Set ChkModel = ActiveSheet.Shapes("ConvChkBx1")
        ChkModel.Copy
        ActiveSheet.Paste Destination:=ConvBKMRK.Offset(-3 + (MbrNMBR + 8) * (i - 1) + 2, 8)
        Set NewChk = ActiveSheet.CheckBoxes(i)
        NewChk.Name = "ConvChkBx" & i
        NewChk.Left = ChkModel.Left
        NewChk.Top = NewChk.Top - 4
        NewChk.LinkedCell = ConvBKMRK.Offset(-3 + (MbrNMBR + 8) * (i - 1) + 2, 8).Address
        
    Next

    Set msgallotrs = MSGraphBKMRK.End(xlToRight).Offset(0, -2)
    Set prsallotrs = prsBKMRK.End(xlToRight)
    
    'Correct msg aotrs formula ttls
    '----------------------------
    aotrsFormula = "=IF($A11=""Q"",$" & msgallotrs.Offset(1, 2).Address(0, 0)
    For i = 1 To suppNMBR + AddNmbr
        aotrsFormula = aotrsFormula & "-$" & msgallotrs.Offset(1, i * -2).Address(0, 0)
    Next
    msgallotrs.Offset(1, 0).Formula = aotrsFormula & ",0)"
    If MbrNMBR > 1 Then msgallotrs.Offset(1, 0).AutoFill Destination:=Range(msgallotrs.Offset(1, 0), msgallotrs.Offset(MbrNMBR, 0))
    
'    Set msgMbrRng = Range(MSGraphBKMRK.Offset(1, 0), MSGraphBKMRK.Offset(MbrNMBR, 0))
'    Set prsMbrRng = Range(prsBKMRK.Offset(1, 0), prsBKMRK.Offset(MbrNMBR, 0))

    'Check to see if suppliers need to be added
    '=======================================================================================
    AddNmbr = 0
    Sheets("line item data").Range("AG:AJ").Calculate
    'Range(MSGraphBKMRK.Offset(MbrNMBR + 8, -1), MSGraphBKMRK.Offset(MbrNMBR * 2 + 8, 7)).Calculate
    Range(MSGraphBKMRK.Offset(1, -1), msgallotrs.Offset(MbrNMBR + 1, 2)).Calculate
    'Range(prsBKMRK, prsallotrs.Offset(MbrNMBR + 2, 0)).Calculate
    If msgallotrs.Offset(MbrNMBR + 1, 1).Value > 0.05 And Not suppNMBR = 10 Then

        'Get Mftr Percentages
        '=======================================================================================
        Call FUN_TestForSheet("xxCalculations")
        Cells.Clear
        For Each c In Range(Sheets("Line item data").Range("V5"), Sheets("Line item data").Range("V" & Sheets("Line item data").Rows.Count).End(xlUp))
            If Not Application.CountIf(Range("A:A"), c.Value) > 0 Then
                CurrRw = Range("A" & Rows.Count).End(xlUp).Row + 1
                Range("A" & CurrRw).Value = c.Value
                Range("B" & CurrRw).Value = WorksheetFunction.SumIfs(Sheets("line item data").Range("AJ:AJ"), Sheets("line item data").Range("V:V"), Range("A" & CurrRw).Value, Sheets("line item data").Range("AI:AI"), "<>x") / Sheets("Line item data").Range("AJ3").Value  '"=SumIf('Line item data'!V:V,A" & CurrRw & ",'Line Item Data'!AJ:AJ)"
                'Range("C" & CurrRw).Value = Range("B" & CurrRw).Value \ Sheets("Line item data").Range("AJ3").Value
            End If
        Next
        On Error Resume Next
        For i = 1 To suppNMBR
            Range("A:A").Find(what:=Sheets("Line Item data").Range("BG3").Offset(0, (i - 1) * 30).Value, lookat:=xlWhole).EntireRow.Delete
        Next
        On Error GoTo 0
        'Range(Range("A1"), Range("C" & Rows.Count).End(xlUp).Row).Calculate
        Call FUN_Sort(ActiveSheet.Name, Range(Range("A1"), Range("B" & Rows.Count).End(xlUp)), Range("B1"), 2)
        
        'Add Suppliers
        '=======================================================================================
        CurrPrct = msgallotrs.Offset(MbrNMBR + 1, 1).Value
        For Each c In Range(Range("A1"), Range("A1").End(xlDown))
            Call AddSupplier_Col(True, c.Value)
            CurrPrct = CurrPrct - Range("B" & c.Row).Value
            AddNmbr = c.Row
            If CurrPrct <= 0.05 Or suppNMBR + c.Row = 10 Then Exit For
        Next
        Application.DisplayAlerts = False
        Sheets("xxCalculations").Delete
    
        'Formatting
        '=======================================================================================
        'prsBKMRK.Offset(MbrNMBR + 1, 1).Copy
        'Range(prsBKMRK.Offset(MbrNMBR + 1, suppNMBR * 2 + 1), prsBKMRK.Offset(MbrNMBR + 1, suppNMBR * 2 + AddNmbr)).PasteSpecial xlPasteFormats
        'Range(prsBKMRK.Offset(1, suppNMBR * 2 + 1), prsBKMRK.Offset(MbrNMBR, suppNMBR * 2 + AddNmbr)).Interior.ColorIndex = 0

    End If

    'Conditional formatting
    '-------------------------
    For i = 1 To suppNMBR + AddNmbr + 1 + 1
        Range(MSGraphBKMRK.Offset(1, i * 2 - 1), MSGraphBKMRK.Offset(MbrNMBR + 1, i * 2 - 1)).NumberFormat = "$#,##0"
        condFormatStr = condFormatStr & Range(MSGraphBKMRK.Offset(1, i * 2 - 1), MSGraphBKMRK.Offset(MbrNMBR + 1, i * 2 - 1)).Address & ", "
    Next
    condFormatStr = Left(condFormatStr, Len(condFormatStr) - 2)
    Sheets("initiative spend overview").Select
    With Range(condFormatStr)
         .FormatConditions.Delete
         .FormatConditions.Add Type:=xlExpression, Formula1:="=$C$8=TRUE"
         .FormatConditions(1).NumberFormat = "0"
    End With

    Application.StatusBar = "Formatting Current Marketshare Graph"
    Range(MSGraphBKMRK, MSGraphBKMRK.End(xlToRight)).WrapText = True
    Range(prsBKMRK, prsBKMRK.End(xlToRight)).WrapText = True

    'Correct msg aotrs formula ttls
    '----------------------------
    aotrsFormula = "=IF($A11=""Q"",$" & msgallotrs.Offset(1, 2).Address(0, 0)
    For i = 1 To suppNMBR + AddNmbr
        aotrsFormula = aotrsFormula & "-$" & msgallotrs.Offset(1, i * -2).Address(0, 0)
    Next
    msgallotrs.Offset(1, 0).Formula = aotrsFormula & ",0)"
    If MbrNMBR > 1 Then msgallotrs.Offset(1, 0).AutoFill Destination:=Range(msgallotrs.Offset(1, 0), msgallotrs.Offset(MbrNMBR, 0))
    Range(msgallotrs.Offset(1, 0), msgallotrs.Offset(MbrNMBR + 1, 2)).Calculate

    'Correct prs aotrs formula ttls
    '----------------------------
    For i = 1 To MbrNMBR
        aotrsStr = ""
        For j = 1 To AddNmbr
            aotrsStr = aotrsStr & "-$" & prsallotrs.Offset(i, -j).Address(0, 0)
        Next
        For j = 1 To suppNMBR
            aotrsStr = aotrsStr & "-$" & prsallotrs.Offset(i, j * -2 - AddNmbr).Address(0, 0)
        Next
        prsallotrs.Offset(i, 0).Formula = Left(prsallotrs.Offset(i, 0).Formula, InStr(prsallotrs.Offset(i, 0).Formula, "<>X") + 5) & aotrsStr & ",0)"
    Next
    Range(prsallotrs.Offset(1, 0), prsallotrs.Offset(MbrNMBR + 2, 0)).Calculate

    'Add new suppliers to current mkt share graph
    '=======================================================================================

    'Set data range and labels
    '----------------------------------
    Sheets("Initiative Spend Overview").Select
    
    For i = 1 To suppNMBR
        axlbl = axlbl & "'Initiative Spend Overview'!" & MSGraphBKMRK.Offset(0, i * 2 - 1).Address & ", "
        grphdata = grphdata & MSGraphBKMRK.Offset(MbrNMBR + 1, i * 2 - 1).Address & ", "
    Next
    For i = 1 To AddNmbr
        axlbl = axlbl & "'Initiative Spend Overview'!" & MSGraphBKMRK.Offset(0, suppNMBR * 2 + i * 2 - 1).Address & ", "
        grphdata = grphdata & MSGraphBKMRK.Offset(MbrNMBR + 1, suppNMBR * 2 + i * 2 - 1).Address & ", "
    Next
    axlbl = axlbl & "'Initiative Spend Overview'!" & msgallotrs.Address
    grphdata = grphdata & msgallotrs.End(xlDown).Address
    
    'Change Axis & Labels
    '----------------------------------
    ActiveSheet.ChartObjects(2).Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SetSourceData Source:=Range(grphdata)
    ActiveChart.SeriesCollection(1).XValues = "=" & axlbl
    ActiveChart.SeriesCollection(1).DataLabels.Select
    'Selection.ShowCategoryName = True
    Selection.ShowValue = True
    Selection.ShowPercentage = True
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    'ActiveChart.Legend.Font.Bold = True
    'ActiveChart.Legend.Font.Size = 12
    ActiveChart.SeriesCollection(1).HasLeaderLines = True
'    'Sheets("initiative spend overview").ChartObjects(2).Chart.SeriesCollection(1).HasDataLabels = True
'    Selection.Position = xlLabelPositionBestFit
'    Sheets("initiative spend overview").ChartObjects(2).SeriesCollection(1).DataLabels.Position = xlLabelPositionBestFit
'    ActiveChart.SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd

    'Format Size & Position
    '----------------------------------
    MSOffset = 0
    Set BenchStart = BenchBKMRK
    WidthFactor = 1
    HeightFactor = 1
    If MbrNMBR < 10 Then
        hghtcells = 5 + MbrNMBR
        MSOffset = -3
        If MbrNMBR < 5 Then
            WidthFactor = 2
        Else
            WidthFactor = 1.5
        End If
        HeightFactor = 1.5
        If prsBKMRK.End(xlToRight).Column > BenchBKMRK.End(xlToRight).Column Then Set BenchStart = prsBKMRK

        'Clear grey Bars
        '----------------------------------
        Set GreyBar = MSGraphBKMRK.End(xlToRight).Offset(MbrNMBR + 3, 2)
        Range(GreyBar, GreyBar.End(xlToRight)).Clear
        
        Set GreyBar = Cells(BenchBKMRK.Offset(MbrNMBR + 3, 0).Row, BenchStart.End(xlToRight).Offset(0, 2).Column)
        Range(GreyBar, GreyBar.End(xlToRight)).Clear
        
    'ElseIf MbrNMBR < 20 Then
        'hghtcells = MbrNMBR + 1
        'WidthFactor = 2
    ElseIf MbrNMBR < 35 Then
        hghtcells = MbrNMBR + 1
        WidthFactor = 1.2
    Else
        hghtcells = 35
    End If
    
    'MSgraph
    '----------------------------------
    Set RngToCover1 = Range(MSGraphBKMRK.End(xlToRight).Offset(MSOffset, 2), MSGraphBKMRK.End(xlToRight).Offset(hghtcells, 2))
    Set ChtOb1 = Sheets("Initiative Spend overview").ChartObjects("Chart 5")
    ChtOb1.Height = RngToCover1.Height '* HeightFactor ' resize
    ChtOb1.Width = ChtOb1.Height * WidthFactor  ' resize
    ChtOb1.Top = RngToCover1.Top       ' reposition
    ChtOb1.Left = RngToCover1.Left
    
    'BenchGraph
    '----------------------------------
    Set RngToCover2 = Range(BenchBKMRK, BenchBKMRK.Offset(hghtcells, 0))
    Set ChtOb2 = Sheets("Initiative Spend overview").ChartObjects("Chart 7")
    ChtOb2.Height = RngToCover2.Height * HeightFactor
    ChtOb2.Width = ChtOb1.Width 'ChtOb2.Height '* WidthFactor
    ChtOb2.Top = RngToCover2.Top
    ChtOb2.Left = BenchStart.End(xlToRight).Offset(0, 2).Left '- BenchBKMRK.End(xlToRight).Offset(0, 2 + BenchOffset - 1).ColumnWidth '/ 2


MSGraphBKMRK.Select
Application.StatusBar = False


End Sub
Sub AddSupplier_Col(NonCon As Boolean, Optional MftrNm As String)

    Set msgallotrs = MSGraphBKMRK.End(xlToRight).Offset(0, -2)
    Set prsallotrs = prsBKMRK.End(xlToRight)
    
    If NonCon = True Then
        PRSoffset = 1
        SuppNm = Chr(34) & MftrNm & Chr(34)
        NmAddress = msgallotrs.Address
    Else
        PRSoffset = 2
        conpos = (MSGraphBKMRK.End(xlToRight).Offset(0, -2).Column - MSGraphBKMRK.Offset(0, 1).Column) / 2 + 1
        SuppNm = "'Index'!" & ConTblBKMRK.Offset(0, conpos).Address
        NmAddress = "'Line Item Data'!" & Range("BG3").Offset(0, (conpos - 1) * 30).Address
    End If
    
    'Insert column
    '----------------------------------
    Range(prsallotrs, prsallotrs.Offset(MbrNMBR + 2, PRSoffset - 1)).Insert Shift:=xlToRight
    Range(msgallotrs, msgallotrs.Offset(MbrNMBR + 1, 1)).Insert Shift:=xlToRight
    Set prsAddCol = prsallotrs.Offset(0, -PRSoffset)
    Set msgAddCol = msgallotrs.Offset(0, -2)
    prsAddCol.Value = "=" & SuppNm
    msgAddCol.Value = "=" & SuppNm
    If NonCon = False Then prsAddCol.Offset(0, 1).Value = "=Concatenate(" & SuppNm & ", "" Reported"")"
    msgAddCol.Offset(0, 1).Value = "%"
    'Range(prsAddCol.Offset(1, 0), prsAddCol.Offset(MbrNMBR, 0)).Interior.ColorIndex = 0
                
    'spend formulas
    '----------------------------------
    'msgaddcol.Offset(1, 0).Formula = "=IF($C$8=FALSE,SUMIFS('Line Item Data'!$AJ:$AJ,'Line Item Data'!$P:$P,$B11,'Line Item Data'!$V:$V," & "" & addcol.Value & "" & "),SUMIFS('Line Item Data'!$AG:$AG,'Line Item Data'!$P:$P,$B11,'Line Item Data'!$V:$V," & "" & addcol.Value & "" & "))"
'    msgAddCol.Offset(1, 0).Formula = "=IF($C$8=FALSE," & prsAddCol.Offset(1, 0).Address(0, 0) & ",SUMIFS('Line Item Data'!$AG:$AG,'Line Item Data'!$P:$P,$B11,'Line Item Data'!$AI:$AI,""<>X"",'Line Item Data'!$V:$V," & NmAddress & "))"
'    prsAddCol.Offset(1, 0).Formula = "=SUMIFS('Line Item Data'!$AJ:$AJ,'Line Item Data'!$P:$P,$B11,'Line Item Data'!$AI:$AI,""<>X"",'Line Item Data'!$V:$V," & NmAddress & ")"
'    If MbrNMBR > 1 Then
'        msgAddCol.Offset(1, 0).AutoFill Destination:=Range(msgAddCol.Offset(1, 0), msgAddCol.Offset(MbrNMBR, 0))
'        prsAddCol.Offset(1, 0).AutoFill Destination:=Range(prsAddCol.Offset(1, 0), prsAddCol.Offset(MbrNMBR, 0))
'    End If
    
    'Spend Totals
    '----------------------------------
    Range(MSGraphBKMRK.Offset(1, 1), MSGraphBKMRK.Offset(MbrNMBR + 1, 2)).Copy
    msgAddCol.Offset(1, 0).PasteSpecial xlPasteAll
    Range(prsBKMRK.Offset(1, 1), prsBKMRK.Offset(MbrNMBR + 2, PRSoffset)).Copy
    prsAddCol.Offset(1, 0).PasteSpecial xlPasteAll
    Range(msgAddCol.Offset(1, 0), msgAddCol.Offset(MbrNMBR, 0)).Replace what:="'Line Item Data'!$BG$3", replacement:=NmAddress, lookat:=xlPart
    Range(prsAddCol.Offset(1, 0), prsAddCol.Offset(MbrNMBR, 0)).Replace what:="'Line Item Data'!$BG$3", replacement:=NmAddress, lookat:=xlPart
    Range(msgAddCol.Offset(1, 0), msgAddCol.Offset(MbrNMBR, 0)).Replace what:="=IF($C$8=FALSE," & Left(msgAddCol.Address(0, 0), 1), replacement:="=IF($C$8=FALSE," & Left(prsAddCol.Address(0, 0), 1), lookat:=xlPart
    If NonCon = False Then Range(prsAddCol.Offset(1, 1), prsAddCol.Offset(MbrNMBR, 1)).ClearContents
    'msgAddCol.Offset(MbrNMBR + 1, 0).Formula = "=SUM(" & msgAddCol.Offset(1, 0).Address(0, 0) & ":" & msgAddCol.Offset(MbrNMBR, 0).Address(0, 0) & ")"
    'prsAddCol.Offset(MbrNMBR + 1, 0).Formula = "=SUM(" & prsAddCol.Offset(1, 0).Address(0, 0) & ":" & prsAddCol.Offset(MbrNMBR, 0).Address(0, 0) & ")"
    
    'percentage formulas for msgraph table
    '----------------------------------
'    msgAddCol.Offset(1, 1).Formula = "=IFERROR(" & msgAddCol.Offset(1, 0).Address(0, 0) & "/$" & MSGraphBKMRK.End(xlToRight).Offset(1, 0).Address(0, 0) & ",0)"
'    If MbrNMBR > 1 Then msgAddCol.Offset(1, 1).AutoFill Destination:=Range(msgAddCol.Offset(1, 1), msgAddCol.Offset(MbrNMBR, 1))
'    msgAddCol.Offset(MbrNMBR + 1, 1).Formula = "=IFERROR(" & msgAddCol.Offset(MbrNMBR + 1, 0).Address(0, 0) & "/$" & MSGraphBKMRK.End(xlToRight).Offset(MbrNMBR + 1, 0).Address(0, 0) & ",0)"
    'prsAddCol.Offset(MbrNMBR + 2, 0).Formula = "=IFERROR(+" & prsAddCol.Offset(MbrNMBR + 1, 0).Address(0, 0) & "/" & prsBKMRK.Offset(-6, 1).Address & ",""-"")"
    Range(msgAddCol, msgAddCol.Offset(MbrNMBR + 1, 1)).Calculate
    Range(prsAddCol, prsAddCol.Offset(MbrNMBR + 2, 0)).Calculate
    
    'Range(prsAddCol.Offset(1, 0), prsAddCol.Offset(MbrNMBR, 0)).Interior.ColorIndex = 0
    
     
End Sub
Sub AddSupplier_Summary(conpos, Bkmrk As Range)

Dim SuppCols(1 To 12) As String
SuppCols(1) = "BG"
SuppCols(2) = "BO"
SuppCols(3) = "BH"
SuppCols(4) = "BP"
SuppCols(5) = "BQ"
SuppCols(6) = "BR"
SuppCols(7) = "BS"
SuppCols(8) = "BT"
SuppCols(9) = "BU"
SuppCols(10) = "BY"
SuppCols(11) = "CD"
SuppCols(12) = "CI"

    'Copy Supplier table
    '----------------------------------------
    Range(Bkmrk.Offset(-3, -1), Bkmrk.Offset(MbrNMBR + 4, 0)).EntireRow.Copy
    Bkmrk.Offset(-3 + (MbrNMBR + 8) * (conpos - 1), -1).PasteSpecial xlPasteAll
        
    'Copy Match Rate Graph
    '----------------------------------------
    Set grphmodel = ActiveSheet.Shapes("MatchGraph1")
    grphmodel.Copy
    ActiveSheet.Paste Destination:=Bkmrk.Offset(-3 + (MbrNMBR + 8) * (conpos - 1), 1)
    Set newGrph = ActiveSheet.ChartObjects(conpos)
    newGrph.Name = "MatchGraph" & conpos
    newGrph.Left = grphmodel.Left
    newGrph.Activate
    ActiveChart.SeriesCollection(1).Values = "='" & ActiveSheet.Name & "'!" & Bkmrk.Offset(-2 + (MbrNMBR + 8) * (conpos - 1), 1).Address
    ActiveChart.SeriesCollection(2).Values = "='" & ActiveSheet.Name & "'!" & Bkmrk.Offset(-2 + (MbrNMBR + 8) * (conpos - 1), 2).Address

    'Replace column references for supplier
    '----------------------------------------
    Set NewRng = Range(Bkmrk.Offset(-3 + (MbrNMBR + 8) * (conpos - 1), -1), Bkmrk.End(xlToRight).Offset(-3 + (MbrNMBR + 8) * (conpos - 1) + MbrNMBR + 3, 0))
    For i = 1 To UBound(SuppCols)
        NewRng.Replace what:="$" & SuppCols(i), replacement:="$" & Right(Range(SuppCols(i) & ":" & SuppCols(i)).Offset(0, (conpos - 1) * 30).Address, 2), lookat:=xlPart
    Next
    Set tierref = Bkmrk.Offset(-2 + (MbrNMBR + 8) * (conpos - 1), 0)
    tierref.Formula = Replace(tierref.Formula, "$D", "$" & Left(ConTblBKMRK.Offset(0, conpos).Address(0, 0), 1))
    Set tierref2 = Bkmrk.Offset((MbrNMBR + 8) * (conpos - 1), 6)                                                        '<--NC only
    tierref2.Formula = Replace(tierref2.Formula, "$D", "$" & Left(ConTblBKMRK.Offset(0, conpos).Address(0, 0), 1))
    Set tierref3 = Bkmrk.Offset((MbrNMBR + 8) * (conpos - 1), 9)                                                        '<--conv only
    tierref3.Formula = Replace(tierref3.Formula, "$D", "$" & Left(ConTblBKMRK.Offset(0, conpos).Address(0, 0), 1))


'    prevref = Left(tierref.Formula, InStr(tierref.Formula, " Pricing'!") - 1)
'    prevref = Mid(prevref, InStrRev(prevref, "'") + 1, Len(prevref))
'    tierref.Formula = Replace(tierref.Formula, prevref & " Pricing", pftab)
    
    'reattach Supplier headers on NC and conv tabs
    '-------------------------------------------------
    pfnm = FUN_SuppName(conpos)
    pftab = "'" & pfnm & " Pricing'!"
    priceformula = Sheets(pfnm & " Pricing").Range("j2").Formula
    If Not priceformula = "" Then
        PFTierRng = Left(priceformula, InStr(priceformula, ")") - 1)
        PFTierRng = Mid(PFTierRng, InStr(PFTierRng, "MIN(") + 4, Len(PFTierRng))
        suppnmAddress1 = "Index!" & ConTblBKMRK.Offset(0, conpos).Address
        suppnmAddress2 = "Index!" & ConTblBKMRK.Offset(1, conpos).Address
        suppnmAddress3 = "Index!" & ConTblBKMRK.Offset(8, conpos).Address
        tierref.Formula = "=CONCATENATE(" & suppnmAddress1 & ","" ""," & suppnmAddress2 & ","" - "",IF(" & suppnmAddress3 & "=""Best Price"",OFFSET(" & pftab & "$J$1,0,MATCH(MIN(" & pftab & PFTierRng & ")," & pftab & PFTierRng & ",0))," & suppnmAddress3 & "))"
        'ConvBKMRK.Offset((MbrNMBR + 8) * (i - 1) - 2, 0).Formula = "=CONCATENATE(" & suppnmAddress1 & ","" "",Index!" & suppnmAddress2 & ","" - "",IF(Index!" & suppnmAddress3 & "=""Best Price"",OFFSET(" & pftab & "$J$1,0,MATCH(MIN(" & pftab & PFTierRng & ")," & pftab & PFTierRng & ",0)),Index!" & suppnmAddress3 & "))"
    End If

    

End Sub
Sub Add_Scenario(connmbr)

    If Sheets("Line item Data").Range("BG4").End(xlToRight).Value = "Owners" Then
        OwnerOffset = 2
    Else
        OwnerOffset = 0
    End If
    conpos = (Sheets("Line item Data").Range("BG4").End(xlToRight).Column - OwnerOffset - Sheets("Line item Data").Range("BG4").Column) / 30
    
    Sheets("line item Data").Range("BG:CJ").Copy
    Sheets("Line item Data").Range("BG1").Offset(0, (conpos - 1) * 30).Insert
    Set NewStart = Sheets("Line item Data").Range("BG1").Offset(0, (conpos - 1) * 30)
    
    If Application.CountIf(ConTblBKMRK.EntireRow, connmbr) > 0 Then
        SuppPos = ConTblBKMRK.Offset(1, 0).EntireRow.Find(what:=connmbr).Column - ConTblBKMRK.Column
        SuppNm = Sheets("Line item Data").Range("BG3").Offset(0, (SuppPos - 1) * 30).Value
        Set pfSht = Sheets(FUN_SuppName(SuppPos) & " Pricing")
        Set maxtier = Range("J1").Offset(0, pfSht.Rows("1:1").Find(what:="EA Price").Column - 11)
        With NewStart.Offset(1, 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='" & pfSht.Name & "'!J1:" & maxtier.Address(0, 0)
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With
        NewStart.Offset(1, 1).Value = "Best Price"
        NewStart.Offset(2, 1).Value = "=IF(" & NewStart.Offset(1, 1).Address & "=""Best Price"",10+MATCH(MIN('" & pfSht.Name & "'!$K$2:" & maxtier.Offset(1, 0).Address & "),'" & pfSht.Name & "'!$K$2:" & maxtier.Offset(1, 0).Address & ",0),10+substitute(" & NewStart.Offset(1, 1).Address & ",""Tier "", ))"
        Range(NewStart.Offset(4, 4), NewStart.Offset(3, 4).End(xlDown)).Replace what:="$A:$J,10", replacement:="$A:$J," & NewStart.Offset(2, 1).Address, lookat:=xlPart
    Else
        SuppNm = "Supplier" & conpos
        Sheets.Add Before:=Sheets("Admin Fees")
        ActiveSheet.Name = SuppNm & " Pricing"
        Sheets.Add Before:=Sheets("Admin Fees")
        ActiveSheet.Name = SuppNm & " Cross Reference"
    End If
    
    Call AddSupplier_Summary(conpos, NonConBKMRK)
    Call AddSupplier_Summary(conpos, ConvBKMRK)

    'Replace Pricefile tab name in formula
    '----------------------------------------
    NewStart.offst(2, 0).Value = SuppNm
    PrevSupp = FUN_SuppName(1)
    Range(NewStart, NewStart.Offset(0, 4)).EntireColumn.Replace what:=PrevSupp, replacement:=SuppNm & " Pricing", lookat:=xlPart


End Sub
Sub ExtractCore()

'PreReq for core xref file setup :
'for each supplier - supplier catnums in one column then assoc descriptions in next column, supplier name at the top of catnum column, no empty columns in between, no extra columns outside of supplier columns. Sheet must be named "Core"
'***************************************************************************************************************

Dim i As Integer

If Trim(ZeusForm.AsscXref.Caption) = "" Then Exit Sub
On Error GoTo ERR_NoXref
Set xrefwb = Workbooks(FUN_OpenWBvar(ZeusPATH, Trim(ZeusForm.AsscXref.Caption), ".xls")) 'Workbooks(FUN_OpenWBvar(ZeusPATH, "CoreXref", pscVar))
On Error GoTo 0

'find last row of max column
'---------------------------------
xrefwb.Activate
On Error GoTo ERR_NoCore
Sheets("DATxref").Visible = True
On Error GoTo 0
Sheets("DATxref").Select
If Trim(LCase(Range("A1").Value)) = "supplier name" Or Trim(Range("A1").Value) = "" Or Trim(Range("A1").Value) = "No Core" Then
    NoCoreFlg = 1
Else
    lastcol = Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    If Not ((lastcol Mod 2) = 0) Then lastcol = lastcol + 1
    LastRow = 0
    For i = 1 To lastcol
        If Cells(ActiveSheet.Rows.Count, i).End(xlUp).Row > LastRow Then LastRow = Cells(ActiveSheet.Rows.Count, i).End(xlUp).Row
    Next
End If

NoCore:
'<LOOP>
'=====================================================================================
For i = 1 To suppNMBR

    DoEvents
    Application.StatusBar = "Importing Core Xref: Supplier(" & i & ")"
    
    'Set supplier name and setup xref tab
    '-------------------------
    tmWB.Activate
    connmbr = ZeusForm.asscContracts.List(i - 1)
    StdName = Sheets("line item data").Range("BG3").Offset(0, (i - 1) * 30).Value
    SuppNm = FUN_SuppName(i)
    Sheets(SuppNm & " Cross Reference").Visible = True
    Sheets(SuppNm & " Cross Reference").Select
    If Trim(Range("C2").Value) = "" Then Range("A2:F2").Value = "x"
    
    If NoCoreFlg = 1 Then GoTo CoreSkip
    
    'set supplier column
    '-------------------------
    xrefwb.Activate
    If Application.CountIf(Rows("1:1"), SuppNm) > 0 Then
        Rows("1:1").Find(what:=SuppNm, lookat:=xlWhole).Select
    ElseIf Application.CountIf(Rows("1:1"), StdNm) > 0 Then
        Rows("1:1").Find(what:=StdNm, lookat:=xlWhole).Select
    Else
        On Error GoTo ERR_NoSupp
        Application.ScreenUpdating = True
        DoEvents
        Set SuppCol = Application.InputBox(prompt:="Please select column containing " & SuppNm & " catalog numbers.  If there is not one for this supplier then cacnel.", Type:=8)
        Application.ScreenUpdating = False
        On Error GoTo 0
    End If
    Set SuppCol = Range(Cells(2, ActiveCell.Column), Cells(LastRow, ActiveCell.Column))
    
    'loop through and copy each supplier xref
    '-------------------------
    For j = 1 To lastcol / 2
        
        Set suppstrt = Range("A1").Offset(0, (j - 1) * 2)
        Set CopyCol = Range(suppstrt.Offset(1, 0), suppstrt.Offset(LastRow, 0))
        If Not CopyCol.Column = SuppCol.Column Then
            xrefNm = suppstrt.Value
            StrtRw = tmWB.Sheets(SuppNm & " Cross Reference").Range("C1").End(xlDown).Offset(1, 0).Row
            tmWB.Sheets(SuppNm & " Cross Reference").Range("A" & StrtRw & ":A" & StrtRw + LastRow - 2).Value = CopyCol.Value
            tmWB.Sheets(SuppNm & " Cross Reference").Range("B" & StrtRw & ":B" & StrtRw + LastRow - 2).Value = SuppCol.Value
            tmWB.Sheets(SuppNm & " Cross Reference").Range("C" & StrtRw & ":C" & StrtRw + LastRow - 2).Value = xrefNm
            tmWB.Sheets(SuppNm & " Cross Reference").Range("D" & StrtRw & ":D" & StrtRw + LastRow - 2).Value = CopyCol.Offset(0, 1).Value
        End If

    Next
    
CoreSkip:
    lastAddedRow = tmWB.Sheets(SuppNm & " Cross Reference").Range("C1").End(xlDown).Row
    tmWB.Sheets(SuppNm & " Cross Reference").Range("F2:F" & lastAddedRow).Value = "Core"
    tmWB.Sheets(SuppNm & " Cross Reference").Range("G2:G" & lastAddedRow).Value = connmbr
    tmWB.Sheets(SuppNm & " Cross Reference").Range("H2:H" & lastAddedRow).Value = ConTblBKMRK.Offset(3, i).Value & " - " & ConTblBKMRK.Offset(4, i).Value

    'import xrefDB if present
    '=============================================================================
    For Each sht In xrefwb.Sheets
        If sht.Name = "XrefDB_" & connmbr Then
            If Not sht.Range("A1").Value = "No xref found" And Not Trim(sht.Range("A1").Value) = "" Then
                lastDBrow = FUN_lastrow("A", sht.Name)
                tmWB.Sheets(SuppNm & " Cross Reference").Range("A" & lastAddedRow + 1 & ":H" & lastrowadded + lastDBrow).Value = sht.Range("A1:H" & lastDBrow).Value
            End If
        End If
    Next
    If Not lastDBrow = "" Then tmWB.Sheets(SuppNm & " Cross Reference").Range("F" & lastAddedRow + 1 & ":F" & lastrowadded + lastDBrow).SpecialCells(xlCellTypeBlanks).Value = "xrefDB"
    
NxtSupp:
If tmWB.Sheets(SuppNm & " Cross Reference").Range("A2").Value = "x" Then tmWB.Sheets(SuppNm & " Cross Reference").Range("A2").EntireRow.Delete Shift:=xlUp
Next

xrefwb.Close (False)
tmWB.Activate
If Trim(Range("C2").Value) = "" Then Range("A2:F2").Value = "x"
Application.StatusBar = False


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_NoXref:
If Not MainCall = 1 Then MsgBox "No core xref with corresponding filename was found."  'Please create your xref file and make sure ""CoreXref"" and the PSC are in the file name then try again."
Exit Sub

ERR_NoCore:
NoCoreFlg = 1
On Error GoTo 0
Resume NoCore
    
ERR_NoSupp:
tmWB.Sheets("notes").Range("F1").Offset(0, i).Value = "No cross reference for " & SuppNm
tmWB.Sheets("notes").Range("F1").Offset(0, i).Font.ColorIndex = 3
On Error GoTo 0
Resume NxtSupp





End Sub
'Sub METH_ExtractIntelli()
'
''Have all intelli contract xref files and xref open, all contracts must be in their own separate file.
''********************************************************************************************************
'Dim fso, oFolder, oSubfolder, ofile, queue As Collection
'
'If Trim(ZeusForm.AsscXref.Caption) = "" Then Exit Sub
'On Error GoTo errhndlNoXref
'Set xrefwb = Workbooks(FUN_OpenWBvar(ZeusPATH, Trim(ZeusForm.AsscXref.Caption), ".xls")) 'Workbooks(FUN_OpenWBvar(ZeusPATH, "CoreXref", pscVar))
'On Error GoTo 0
'
'
''Close each intelli xref if open, then open all that are in gentry folder and execute
''==================================================================================================
'For Each Wb In Workbooks
'    If InStr(Wb.Name, "Contract_Cross_Reference_Report") > 0 Or InStr(Wb.Name, "Contract Cross Reference Report") > 0 Then
'        Wb.Close (False)
'    End If
'Next
'
'setobjFolder (ZeusPATH)
'
''execute
''==================================================================================================
'For Each ofile In objFolder.Files
'    If InStr(ofile.Name, "Contract_Cross_Reference_Report") > 0 Or InStr(ofile.Name, "Contract Cross Reference Report") > 0 Then
'        Set intelliWB = Workbooks.Open(ofile.Path)
'        If IsEmpty(Range("E4")) Then GoTo SAMENAME
'        ConNmbr = Range("A4").Value
'
'        'find suppNm
'        '------------------------------
'        tmWB.Activate
'        Sheets("impact summary").Select
'        Cells.Find(what:="Agreement Number", After:=ActiveCell, LookIn:=xlFormulas, _
'            lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
'            MatchCase:=False, SearchFormat:=False).Select
'        Range(ActiveCell.Offset(0, 1), ActiveCell.End(xlToRight)).Select
'        Selection.Find(what:=ConNmbr, After:=ActiveCell, LookIn:=xlFormulas, _
'            lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
'            MatchCase:=False, SearchFormat:=False).Select
'        nmrng = Application.CountA(Range(ActiveCell, ActiveCell.End(xlToLeft))) - 1
'        SuppNm = Sheets("notes").Range("AA1").Offset(nmrng, 0).Value
'        If IsEmpty(SuppNm) Then
'            SuppNm = ActiveCell.Offset(-1, 0).Value
'        End If
'        DoEvents
'        Application.StatusBar = "Importing Intellisource for " & SuppNm & ")"
'
'        'pull in cross
'        '------------------------------
'        shtflg = 0
'        For Each sht In tmWB.Sheets
'            If InStr(sht.Name, SuppNm & " Cross Reference") > 0 Then
'1               shtflg = 1
'                sht.Visible = True
'                sht.Select
'                If Not IsEmpty(Range("F2")) Then
'
'                    'contract catnum
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("E4"), Range("E3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("F1").End(xlDown).Offset(1, -4).Select
'                    ActiveSheet.Paste
'
'                    'cross catnum
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("K4"), Range("K3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("F1").End(xlDown).Offset(1, -5).Select
'                    ActiveSheet.Paste
'
'                    'cross desc
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("L4"), Range("L3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("F1").End(xlDown).Offset(1, -2).Select
'                    ActiveSheet.Paste
'
'                    'cross suppnm
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("I4"), Range("I3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("F1").End(xlDown).Offset(1, -3).Select
'                    ActiveSheet.Paste
'
'                    'source
'                    '--------------------
'                    Range(Range("F1").End(xlDown).Offset(1, 0), Range("C1").End(xlDown).Offset(0, 3)).Value = "Intellisource"
'
'                    Exit For  'dont go through the rest of the sheets
'
'                Else
'                    'contract CatNum
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("E4"), Range("E3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("B2").Select
'                    ActiveSheet.Paste
'
'                    'cross catnum
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("K4"), Range("K3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("A2").Select
'                    ActiveSheet.Paste
'
'                    'cross desc
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("L4"), Range("L3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("D2").Select
'                    ActiveSheet.Paste
'
'                    'cross suppnm
'                    '--------------------
'                    intelliWB.Activate
'                    Range(Range("I4"), Range("I3").End(xlDown)).Copy
'                    tmWB.Activate
'                    Range("C2").Select
'                    ActiveSheet.Paste
'
'                    'source
'                    '--------------------
'                    Range(Range("F2"), Range("A1").End(xlDown).Offset(0, 5)).Value = "Intellisource"
'                    Range(Range("G2"), Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Offset(0, 4)).Value = Sheets("impact summary").Range("A:A").Find(what:="Supplier name").Offset(1, suppcnt).Value
'                    Range(Range("H2"), Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Offset(0, 5)).Value = Sheets("impact summary").Range("A:A").Find(what:="Supplier name").Offset(4, suppcnt).Value
'
'                    Exit For    'dont go through the rest of the sheets
'
'                End If
'            End If
'        Next
'
'        'if looped through all shts and not found then manually input suppnm
'        '---------------------
'        If shtflg = 0 Then
'            handlecnt1 = 0
'            On Error GoTo Errhandler3
'            Sheets(SuppNm & " Cross Reference").Visible = True
'            Sheets(SuppNm & " Cross Reference").Select
'            GoTo 1
'        End If
'
'        'copy to xref
'        '---------------------
'        intelliWB.Activate
'        On Error GoTo errhndlSAMENAME
'        ActiveSheet.Copy After:=xrefwb.Sheets("DATxref")
'        xrefwb.Activate
'        Application.DisplayAlerts = False
'        ActiveSheet.Name = SuppNm
'        Application.DisplayAlerts = True
'SAMENAME:
'        On Error GoTo 0
'        intelliWB.Close (False)
'    End If
'
'Next
'
'Application.StatusBar = False
'createxref: Exit Sub
'':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'errhndlNoXref:
'If Not MainCall = 1 Then MsgBox "No core xref was found.  Please create your xref file and make sure ""CoreXref"" and the PSC are in the file name then try again."
'Exit Sub
'
'Errhandler3:
'    'Manually input the supplier name, if can't get it to match after 3 tries then turn off error handler and debug
'    '---------------------------------
'    handlecnt1 = handlecnt1 + 1
'
'    If handlecnt1 < 4 Then
'       SuppNm = Application.InputBox(prompt:="Supplier name on Impact Summary does not match Supplier name associated with pricefile.  Please enter supplier name associated with the pricefile tab", Title:="Supplier Mismatch", Type:=2)
'    Else
'       MsgBox ("Please check supplier name and try again")
'       On Error GoTo 0
'    End If
'Resume Next
'
'errhndlSAMENAME:
'Resume SAMENAME
'
'
'End Sub
Sub FormatCrossRefTabs()
    
    
    For i = 1 To suppNMBR
    
        DoEvents
        Application.StatusBar = "Formatting Xref tabs: Supplier(" & i & ")"

        'Set supplier name and goto pricefile
        '-------------------------
        SuppNm = FUN_SuppName(i)
        Sheets(SuppNm & " Cross Reference").Visible = True
        Sheets(SuppNm & " Cross Reference").Select
        
        'Cross ref setup
        '======================================================================================
        If Application.CountA(Range("A:A")) - 1 = 0 Then GoTo NxtSupp
        
        'convert to number
        '----------------------------
        On Error GoTo 0
'        Columns("A:A").Select
'        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
'            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'            :=Array(1, 1), TrailingMinusNumbers:=True
'        Columns("B:B").Select
'        Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
'            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
'            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
'            :=Array(1, 1), TrailingMinusNumbers:=True

        'format headers
        '----------------------------
        Rows("1:1").Clear
        Range("A1").Value = "Member SKU"
        Range("B1").Value = "Contract SKU"
        Range("C1").Value = "Supplier"
        Range("D1").Value = "Product Description"
        Range("E1").Value = UCase(SuppNm & " description")
        Range("F1").Value = "Source"
        Range("G1").Value = "Contract"
        Range("H1").Value = "Date"
        Range("I1").Value = UCase(SuppNm & " EA price")
        
        'Vlookup desc and EA price
        '-----------------------------
        Range("E2").Formula = "=vlookup(B2,'" & SuppNm & " Pricing'!$A:$CA,3,false)"
        Range("E2").AutoFill Destination:=Range("E2:E" & Cells(Rows.Count, "F").End(xlUp).Row)
        Range("E2:E" & Cells(Rows.Count, "F").End(xlUp).Row).Calculate
        On Error GoTo ERR_EAcol
        EAcol = Sheets(SuppNm & " Pricing").Rows("1:1").Find(what:="Ea", lookat:=xlPart).Column
        On Error GoTo 0
        Range("I2").Formula = "=vlookup(B2,'" & SuppNm & " Pricing'!$A:$CA," & EAcol & ",false)"
        Range("I2").AutoFill Destination:=Range("I2:I" & Cells(Rows.Count, "F").End(xlUp).Row)
        Range("I2:I" & Cells(Rows.Count, "F").End(xlUp).Row).Calculate
        Range("I:I").NumberFormat = "$#,##0"
        
        'Remove N/As
        '------------------------------
        ActiveWorkbook.Worksheets(SuppNm & " Cross Reference").Sort.SortFields.Clear
        Call FUN_Sort(SuppNm & " Cross Reference", Range("A2:I200000"), Range("I2"), 1)
        Range(Range("I2"), Range("F1").End(xlDown).Offset(0, 3)).Copy
        Range("O2").PasteSpecial Paste:=xlPasteValues
        On Error GoTo ERR_NoNA
        Range("O:O").Find(what:="#N/A", lookat:=xlWhole).Select
        On Error GoTo 0
        Application.StatusBar = "Removing #N/A..."
        Range(ActiveCell, ActiveCell.End(xlDown)).EntireRow.Delete Shift:=xlUp
        
        'Sort
        '------------------------------
NoNA:   Application.StatusBar = "Sorting for Core and lowest price..."
        ActiveWorkbook.Worksheets(SuppNm & " Cross Reference").Sort.SortFields.Clear
        Call FUN_Sort(SuppNm & " Cross Reference", Range("A2:I200000"), Range("A2"), 1, Range("F2"), 1, Range("I2"), 1)
        
        'delete empties in col A
        '-------------------------------
        Application.StatusBar = "Removing #N/A..."
        On Error GoTo ERR_NoXref
        Range(Range("A1").End(xlDown).Offset(1, 0), Range("F1").End(xlDown).Offset(1, 0)).EntireRow.Delete Shift:=xlUp
        On Error GoTo 0
        Range(Range("O:O"), Range("O:O").End(xlToRight)).Clear

clean:  Application.StatusBar = False
        Range("A:H").HorizontalAlignment = xlLeft
        Range("A1:H1").HorizontalAlignment = xlCenter
        Range("A:O").Borders.LineStyle = xlNone
        Range("A1:I1").Borders.LineStyle = xlContinuous
        Range("A:C").EntireColumn.ColumnWidth = 20
        Range("D:E").EntireColumn.ColumnWidth = 40
        Range("F:H").ColumnWidth = 12.75
        Range("I:I").ColumnWidth = 15
        Range("A:O").Interior.ColorIndex = 0
        Range("A1:I1").Interior.ColorIndex = 15
        Range("A:I").WrapText = False
NxtSupp:
    Next

Application.StatusBar = False
Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_EAcol:
Sheets(SuppNm & " Pricing").Visible
Sheets(SuppNm & " Pricing").Select
Set eacolrng = Application.InputBox(prompt:="Standard each pricing header not found, please select Each Price Column.", Type:=8)
EAcol = eacolrng.Column
Sheets(SuppNm & " Cross Reference").Select
Resume Next

ERR_NoNA:
On Error GoTo 0
Resume NoNA

ERR_NoXref:
Range(Range("B2"), Range("F1").End(xlDown)).EntireRow.Clear
Range("A2").Value = "No Core or Intellisource xref"
Range("A2").Interior.ColorIndex = 3
On Error GoTo 0
Resume clean


End Sub
Sub METH_UploadXref()

Dim adoRecSet As New ADODB.Recordset
Dim connDB As New ADODB.Connection
Dim strDB As String
Dim strSQL As String
Dim colvar As String

'        strDB = "C:\Users\bforrest\Desktop\Zeus\1-DB shortcuts\Product Segmentation.accdb"
'        connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
'        'Set adoRecSet = New ADODB.Recordset
'        For Each c In Range(Range("A1"), Range("A1").End(xlDown))
'            Debug.Print c.Value
'            connDB.Execute "INSERT INTO [Table1] ([connmbr]) VALUES (" & "'" & c.Value & "'" & ")"
'        Next
'        'connDB.Execute "INSERT INTO [Table1] ([connmbr]) SELECT * FROM [Excel 8.0;HDR=YES;DATABASE= C:\Users\bforrest\Desktop\Zeus\Working File(PAIN PUMPS).xlsx].[Sheet1$]"
'        connDB.Close


    If Not CreateReport = True Then
        xrefCheck.Show
        If endFLG = 1 Then
            endFLG = 0
            Exit Sub
        End If
    End If

    loopcount = 0
    Do
    loopcount = loopcount + 1
        
        'Set supplier name and goto xref tab
        '-----------------------------------
        SuppNm = Sheets("Notes").Range("AA1").Offset(loopcount, 0).Value
        Worksheets(SuppNm & " Cross Reference").Visible = True
        Sheets(SuppNm & " Cross Reference").Select
        
        'check to make sure formatted correctly
        '=============================================================================================
'        Range("A:H").EntireColumn.AutoFit
'        Range("A:H").Copy
'        Range("P:W").PasteSpecial xlPasteValues
        
        'Check to see if there's any items to upload
        '-----------------------------------
        If Not Application.CountA(Range("A:H")) > 8 Then
            GoTo NoUpload
        End If
        
        If Not (CreateReport = True Or DragonSlayerFLG = 1) Then

            'Check Date column
            '-----------------------------------
            If Application.CountA(Range("H:H")) > 1 Then
                If Not IsDate(Range("H1").End(xlDown).Value) Then
                    Range("O1").Value = "Formatting error. Column H must be in date format"
                    Range("O1").Interior.ColorIndex = 3
                    GoTo NoUpload
                End If
            End If
            
            'Check text columns
            '-----------------------------------
            For i = 1 To 7
                
                'check if the column has any data
                '-----------------------------------
                If Application.CountA(Range("A:A").Offset(0, i)) > 1 Then
                    colvar = Left(Range("A1").Offset(0, i).Address(0, 0), 1)
                    
                    '[TBD]is this needed? can either text of number be uploaded in field formatted for text?
                    'check each cell in column to make sure all are text/general
                    '-----------------------------------
                    For Each c In Range(Range("A2").Offset(0, i), Range("A2").Offset(FUN_lastrow(colvar), i))
                        If Not c.Value Is Text Then
                            Range("O1").Value = "Formatting error. Cell " & c.Address(0, 0) & " must be in text/general format"
                            Range("O1").Interior.ColorIndex = 3
                            GoTo NoUpload
                        End If
                    Next
                End If
            Next
            
        End If
        Range("O1").Clear
        Application.StatusBar = "Uploading to xref database.  Please Wait..."

        'upload to DB
        '-----------------------------------
        On Error GoTo errhndlDBERROR
        If Not preconnectFLG = 1 Then
            If LCase(Usr) = "jokiri" Then
                strDB = "Z:\Analytics\DAT Resources\Xref\Xref.accdb"
            Else
                strDB = "\\filecluster01\dfs\NovSecure2\SupplyNetworks\Analytics\DAT Resources\Xref\Xref.accdb"
            End If
            connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
            preconnectFLG = 1
        End If
                    
        'loop through each row and insert
        '-----------------------------------
        ttlupload = Range(Range("A2"), Range("A2").Offset(FUN_lastrow("G"), 0)).Count
        For Each c In Range(Range("A2"), Range("A2").Offset(FUN_lastrow("G"), 0))
'            Debug.Print "INSERT INTO [Xref Table] ([Member Catalog Number] , [Contract Catalog Number], [Cross Reference Supplier], [Member Item Description], [Contract Item Description], Source, Contract, Date) VALUES ('" & c.Value & "','" & c.Offset(0, 1).Value & "','" & c.Offset(0, 2).Value & "','" & c.Offset(0, 3).Value & "','" & c.Offset(0, 4).Value & "','" & c.Offset(0, 5).Value & "','" & c.Offset(0, 6).Value & "','" & c.Offset(0, 7).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] ([Member Catalog Number]) VALUES ('" & c.Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] ([Contract Catalog Number]) VALUES ('" & c.Offset(0, 1).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] ([Cross Reference Supplier]) VALUES ('" & c.Offset(0, 2).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] ([Member Item Description]) VALUES ('" & c.Offset(0, 3).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] ([Contract Item Description]) VALUES ('" & c.Offset(0, 4).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] (Source) VALUES ('" & c.Offset(0, 5).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] (Contract) VALUES ('" & c.Offset(0, 6).Value & "')"
'            connDB.Execute "INSERT INTO [Xref Table] ([Date]) VALUES ('" & c.Offset(0, 7).Value & "')"
            Application.StatusBar = "Uploading Row: (" & c.Row & ") " & (c.Row \ ttlupload) * 100 & "%"
            connDB.Execute "INSERT INTO [Xref Table] ([Member Catalog Number] , [Contract Catalog Number], [Cross Reference Supplier], [Member Item Description], [Contract Item Description], Source, Contract, [Date]) VALUES ('" & c.Value & "','" & c.Offset(0, 1).Value & "','" & c.Offset(0, 2).Value & "','" & c.Offset(0, 3).Value & "','" & c.Offset(0, 4).Value & "','" & c.Offset(0, 5).Value & "','" & c.Offset(0, 6).Value & "','" & c.Offset(0, 7).Value & "')"
        Next
        
        '[TBD] maybe try to format xref as table and use select all method
        'connDB.Execute "INSERT INTO [Xref Table] ([Member Catalog Number] , [Contract Catalog Number], [Cross Reference Supplier], [Member Item Description], [Contract Item Description], Source, Contract, Date) SELECT * FROM [Excel 8.0;HDR=YES;DATABASE=" & ActiveWorkbook.Path & "]." & ActiveSheet.Name
        
        On Error GoTo 0
        Range("O1").Value = "Uploaded Successfuly"
        Range("O1").Interior.ColorIndex = 10

NoUpload: Loop Until loopcount = suppNMBR

connDB.Close
Application.StatusBar = False

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlDBERROR:
Range("O1").Value = "Error connecting to database."
Range("O1").Interior.ColorIndex = 3
On Error GoTo 0
Exit Sub


End Sub
Sub METH_Finalize()

svnm = Replace(ActiveWorkbook.Name, ".xlsx", "")
svnm = Replace(svnm, "(PreQC)", "")
svnm = Replace(svnm, "(POST QC)", "")
svnm = Replace(svnm, "(Final)", "") & "(Final).xlsx"

Application.DisplayAlerts = False
ChDir ZeusPATH
ActiveWorkbook.SaveAs Filename:=ZeusPATH & svnm, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = True


'[TBD]Post to Cache
'====================================================================================================================================
'(post to respective folder, include date, analyst usr)

'UOMs
'--------------------
'open UomdirPATH
'C/P UOMs

tmWB.Close (False)

'contract folder
'--------------------
'If Dir(ContractDirPATH & ConNmbr, vbDirectory) = vbNullString Then
    'if no matching contract then create folder for it
    'copy coreXref,
    'copy pricefiles
'End If

'PSC folder
'--------------------
'If Dir(pscDirPATH & pscVar, vbDirectory) = vbNullString Then
    'if no matching contract then create folder for it
    'end if
    'copy DATxref
    'copy WF
    'copy keywords list
    'FileCopy ZeusPATH & svnm, pscDirPATH & pscVar & "\" & svnm   '<--copy report
'End If


'statsTracker
'====================================================================================================================================
'record stats
'[TBD] input box, suggestions? send metrics email with suggestion

'Post to network folder
'====================================================================================================================================
If Dir(NetworkPath & FileName_PSC, vbDirectory) = vbNullString Then
    MsgBox "No " & "" & FileName_PSC & "" & " folder found for this PSC.  Please submit your report manually."
    retval = Shell("explorer.exe " & NetworkPath, vbNormalFocus)
Else
    If Dir(NetworkPath & FileName_PSC & "\1 DIR Files", vbDirectory) = vbNullString Then
        MsgBox "No ""1 DIR Files"" folder found for this PSC.  Please submit your report manually."
        retval = Shell("explorer.exe " & NetworkPath & FileName_PSC, vbNormalFocus)
    Else
        If Dir(NetworkPath & FileName_PSC & "\1 DIR Files\4 Results (DIR; Market Share)", vbDirectory) = vbNullString Then
            MsgBox "No ""4 Results (DIR; Market Share)"" folder found for this PSC.  Please submit your report manually."
            retval = Shell("explorer.exe " & NetworkPath & FileName_PSC & "\1 DIR Files\", vbNormalFocus)
        Else
            FileCopy ZeusPATH & svnm, NetworkPath & FileName_PSC & "\1 DIR Files\4 Results (DIR; Market Share)\" & svnm
            retval = Shell("explorer.exe " & NetworkPath & FileName_PSC & "\1 DIR Files\4 Results (DIR; Market Share)\", vbNormalFocus)
        End If
    End If
End If




End Sub


