Attribute VB_Name = "ToolBox"
Dim switchTripPrev
Dim switchVal As Integer
Dim FillSwitchTripPrev
Dim FillSwitchVal As Integer
Sub Email_Image()


    'email without opening new message
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    'On Error Resume Next
    With OutMail
        .To = "barry.forrest@vizientinc.com"
        .Subject = "test"
        .Attachments.Add "C:\SavedRange.jpg", olByValue, 0  '<--0 makes the attachment hidden
        .HTMLBody = "<img src='SavedRange.jpg'>"
        .send   'or use .Display
    End With



End Sub
Sub PricefileAdjustments_Indiv()
Attribute PricefileAdjustments_Indiv.VB_ProcData.VB_Invoke_Func = "P\n14"

'Shortcut = Ctrl+Shift+P
'************************
    
    On Error GoTo errhndlCantFind
    Set FirstCat = Range("N:N").Find(what:=ActiveCell.Offset(0, Range("N1").Column - ActiveCell.Column).Value, lookat:=xlWhole)
    mfgname = ActiveCell.Offset(0, Range("L1").Column - ActiveCell.Column).Value
    mfgcol = Rows("1:1").Find(what:=mfgname, lookat:=xlWhole).Address(0, 0)
    mfgcatnmbr = Range(Left(mfgcol, 2) & ActiveCell.Row).Offset(0, 1).Value
    mfgpos = (Range(mfgcol).Column - Range("U1").Column) / 18
    shtnm = Sheets("Notes").Range("AA1").Offset(mfgpos, 0).Value

    Sheets(shtnm & " Pricing").Range("A:A").Find(what:=mfgcatnmbr, lookat:=xlWhole).Offset(0, 4).Value = Application.InputBox(prompt:="Please enter UOM correction:", Type:=1)
    
    mtchcats = Application.CountIf(Range("N:N"), FirstCat.Value)
    Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0)).EntireRow.Calculate
    ActiveCell.EntireRow.Calculate
    
Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlCantFind:
MsgBox ("Couldn't find manufacturer catalog number in pricefile.")
Resume
Exit Sub


End Sub
Sub HighlightDups()

Dim DupCol As String

'Ctrl+Shift+D
'***********************
If Selection.Address = Selection.EntireColumn.Address Then
    DupCol = Replace(FUN_AlphaOnly(ActiveCell.Address), " ", "")
    Set duprng = Range(ActiveCell, ActiveCell.Offset(FUN_lastrow(DupCol) - 1, 0))
ElseIf Selection.Address = Selection.EntireRow.Address Then
    Set duprng = Range(ActiveCell, ActiveCell.Offset(0, FUN_lastcol(ActiveCell.Row) - 1))
Else
    Set duprng = Selection
End If

    For Each c In duprng
        If Application.CountIf(Selection, c.Value) > 1 Then c.Interior.Color = 16711935
    Next

End Sub
Sub Uppercase()

'make all uppercase letters
'Ctrl + Shift + U
'****************************

Dim UppTxt As Variant

For Each UppTxt In Selection
    UppTxt.Value = UCase(UppTxt.Value)
Next


End Sub
Sub TxtToCol()

'Ctrl+Shift+T
'****************************

    For Each c In Selection
        destRng = c.Address
        Exit For
    Next

    Selection.TextToColumns Destination:=Range(destRng), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

End Sub
Sub AddBorders()
Attribute AddBorders.VB_ProcData.VB_Invoke_Func = "B\n14"

'shortcut = Shift+B x1=outside,x2=all,x3=none (w/n 2 sec)
'****************************
Dim switchDif As Long

If Not IsEmpty(switchTripPrev) Then
    switchTrip = Time
    switchDif = DateDiff("s", switchTrip, switchTripPrev)
    If Abs(switchDif) < 2 Then
        switchVal = switchVal + 1
        If switchVal = 4 Then switchVal = 1
    Else
        switchVal = 1
    End If
Else
    switchVal = 1
End If

If switchVal = 1 Then
    Selection.BorderAround ColorIndex:=1 ', Weight:=xlThick
ElseIf switchVal = 2 Then
    Selection.Borders.LineStyle = xlContinuous
Else
    Selection.Borders.LineStyle = xlNone
End If

switchTripPrev = Time



End Sub
Sub calcSelection()
Attribute calcSelection.VB_ProcData.VB_Invoke_Func = "Q\n14"

'shortcut = Ctrl+Shift+Q
'****************************
Selection.Calculate


End Sub
Sub RemoveIOC()
Attribute RemoveIOC.VB_ProcData.VB_Invoke_Func = "R\n14"

'shortcut= Shift+R
'****************************

If Not ActiveSheet.Name = "Line Item Data" Then
    MsgBox "Please select items on Line Item Data tab"
    Exit Sub
End If

Call FUN_CalcOff

If Worksheets("Line Item Data").FilterMode = True Then
    Range("A1:OZ1").EntireColumn.Hidden = False
    Selection.SpecialCells(xlCellTypeVisible).Select
End If
    
Selection.EntireRow.Copy
Sheets("items removed").Visible = True
Sheets("Items Removed").Select
Range("A1").Offset(FUN_lastrow("X"), 0).Select
ActiveSheet.Paste
Sheets("Line Item Data").Select
Selection.EntireRow.Delete Shift:=xlUp
    
Call FUN_CalcBackOn

End Sub
Sub addIOC()

'shortcut= Shift+A
'****************************

If Not ActiveSheet.Name = "Items Removed" Then
    MsgBox "Please select items on Items Removed tab"
    Exit Sub
End If

Call FUN_CalcOff

rws = Selection.Rows.Count

If Worksheets("Items Removed").FilterMode = True Then
    Range("A1:OZ1").EntireColumn.Hidden = False
    Selection.SpecialCells(xlCellTypeVisible).Select
End If
    
Selection.EntireRow.Copy
Sheets("Line Item Data").Select
Range("A6").Insert Shift:=xlDown
Rows("5:5").Copy
Rows("6:" & 5 + rws).PasteSpecial xlPasteFormats

Range("Z5").AutoFill Destination:=Range(Range("Z5"), Range("Z5").Offset(rws, 0))
Range("AG5:AH5").AutoFill Destination:=Range(Range("AG5"), Range("AH5").Offset(rws, 0))
Range("AJ5").AutoFill Destination:=Range(Range("AJ5"), Range("AJ5").Offset(rws, 0))
Range(Range("AN5"), Range("BG5").Offset(0, (suppNMBR - 1) * 30)).AutoFill Destination:=Range(Range("AN5"), Range("BG5").Offset(rws, (suppNMBR - 1) * 30))

Range(Range("A6"), Range("A5").Offset(rws, 0)).EntireRow.Calculate
Sheets("Items Removed").Select
Selection.EntireRow.Delete Shift:=xlUp
    
Call FUN_CalcBackOn

End Sub
Sub DeleteLeft()
Attribute DeleteLeft.VB_ProcData.VB_Invoke_Func = "D\n14"

'shortcut= Shift+C
'****************************

Selection.Delete Shift:=xlToLeft

End Sub
Sub DeleteUp()
Attribute DeleteUp.VB_ProcData.VB_Invoke_Func = "E\n14"

'shortcut= Shift+D
'****************************

Selection.Delete Shift:=xlUp

End Sub
Sub NormalizeData()

'Replaces:   ,.-/(spaces)
    Selection.Replace what:=" ", replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=",", replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="-", replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="/", replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
End Sub
Sub CPvalues()
Attribute CPvalues.VB_ProcData.VB_Invoke_Func = "V\n14"

'shortcut= Shift+V
'****************************

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
Sub AddHighlight()
Attribute AddHighlight.VB_ProcData.VB_Invoke_Func = "A\n14"

'Shortcut= Shift+A  x1=Fill,x2=NoFill (w/n 1 sec)
'****************************
Dim switchDif As Long

If Not IsEmpty(FillSwitchTripPrev) Then
    switchTrip = Time
    switchDif = DateDiff("s", switchTrip, FillSwitchTripPrev)
    If Abs(switchDif) < 1 Then
        If FillSwitchVal = 1 Then
            FillSwitchVal = 2
            Selection.Interior.ColorIndex = 0
        Else
            FillSwitchVal = 1
            On Error Resume Next
            Application.CommandBars.ExecuteMso "CellFillColorPicker"
        End If
    Else
        FillSwitchVal = 1
        On Error Resume Next
        Application.CommandBars.ExecuteMso "CellFillColorPicker"
    End If
Else
    FillSwitchVal = 1
    On Error GoTo sheetlocked
    Application.CommandBars.ExecuteMso "CellFillColorPicker"
End If

FillSwitchTripPrev = Time

Exit Sub
':::::::::::::::::::::::::::::::::::::
sheetlocked:
If InStr(ActiveSheet.Name, "QC") Then MsgBox "Selection is locked"

    
End Sub
Sub xxSelectVisible()
Attribute xxSelectVisible.VB_ProcData.VB_Invoke_Func = "X\n14"
'shortcut= Shift+X

Selection.SpecialCells(xlCellTypeVisible).Select

End Sub
Sub ReplaceNA()
    
    Selection.Replace what:="#N/A", replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
End Sub
Sub NewInstance()

    Dim currentExcel As Workbook
    Dim newExcel As Excel.Application

    Set currentExcel = ActiveWorkbook
    Set newExcel = CreateObject("excel.application")

    newExcel.Visible = True
    newExcel.Workbooks.Add

End Sub
Sub HighlightCatNmbrs()

'Highlight alternating mtchcats for better visibility
'********************************************************8

Range("N3").Select
mtchcats = Application.CountIf(Range(Range("N3"), Range("N3").End(xlDown)), ActiveCell.Value)
Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(mtchcats - 1, 3)).Interior.ColorIndex = 37
ActiveCell.Offset(mtchcats, 0).Select

Do
    mtchcats = Application.CountIf(Range(Range("N3"), Range("N3").End(xlDown)), ActiveCell.Value)
    If ActiveCell.Offset(-1, 3).Interior.Color = 652801 Then
        Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(mtchcats - 1, 3)).Interior.ColorIndex = 37
    Else
        Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(mtchcats - 1, 3)).Interior.Color = 652801
    End If
    ActiveCell.Offset(mtchcats, 0).Select
Loop Until Trim(ActiveCell.Value) = ""


End Sub
Sub RemoveAltPriceProg()

'must be sorted on col A first
'***********************

For Each c In Range(Range("A2"), Range("A2").End(xlDown))
    mtchcats = Application.CountIf(Range("A:A"), c.Value)
    If mtchcats > 1 Then
        Range(c.Offset(1, 0), c.Offset(mtchcats - 1, 0)).Copy
        c.Offset(1, 27).PasteSpecial xlPasteAll
        Range(c.Offset(1, 0), c.Offset(mtchcats - 1, 0)).ClearContents
    End If
Next


End Sub
Sub goTMwb()


For Each Wb In Workbooks
    If InStr(UCase(Wb.Name), UCase(FileName_PSC)) > 0 And InStr(UCase(Wb.Name), UCase(NetNm)) > 0 Then
        Wb.Activate
        Exit Sub
    End If
Next

MsgBox "No Tier Max workbook matching network and PSC"



End Sub
Sub GoNotes()

On Error GoTo NoNotesTab
'On Error GoTo 0
Sheets("Notes").Visible = True
Sheets("notes").Select

Exit Sub
'::::::::::::::::::::::::::::
NoNotesTab:
MsgBox "No Notes tab found"
Exit Sub


End Sub
Sub GoQC() '(if exists)

For Each sht In ActiveWorkbook.Sheets
    If InStr(sht.Name, "QC") Then
        sht.Visible = True
        sht.Select
        Exit Sub
    End If
Next


End Sub
Sub GoIndex()

On Error GoTo NoIdxTab
Sheets("Index").Visible
Sheets("Index").Select

Exit Sub
'::::::::::::::::::::::::::::
NoIdxTab:
MsgBox "No Index tab found"
Exit Sub


End Sub
Sub GoCMSg()

On Error GoTo NoCMSgTab
Sheets("Current Market Share").Visible
Sheets("Current Market Share").Select

Exit Sub
'::::::::::::::::::::::::::::
NoCMSgTab:
MsgBox "No Current Market Share tab found"
Exit Sub


End Sub
Sub GoLID()


On Error GoTo NoLIDTab
Sheets("Line Item Data").Visible
Sheets("Line Item Data").Select

Exit Sub
'::::::::::::::::::::::::::::
NoLIDTab:
MsgBox "No Line Item Data tab found"
Exit Sub


End Sub
Sub GoBMP()

On Error GoTo NoBMPTab
Sheets("Best Market Price").Visible
Sheets("Best Market Price").Select

Exit Sub
'::::::::::::::::::::::::::::
NoBMPTab:
MsgBox "No Best Market Price tab found"
Exit Sub


End Sub
Sub GoItemsRemoved()

On Error GoTo NoIRTab
Sheets("Items Removed").Visible
Sheets("Items Removed").Select

Exit Sub
'::::::::::::::::::::::::::::
NoIRTab:
MsgBox "No Items Removed tab found"
Exit Sub


End Sub
Sub TabNext()


'If ActiveSheet.Name = "Line Item Data" Then
    Sheets("Line Item Data").Visible = True
    Sheets("Line Item Data").Select
'elseif activesheet.name = "



End Sub
Sub TabPrev()


'If ActiveSheet.Name = "Line Item Data" Then
    Sheets("Impact Summary").Visible = True
    Sheets("Impact Summary").Select
'elseif activesheet.name = "



End Sub
Sub AddNewItems()

'from TMwb, have extract open, DIRT ASF in filename
'*******************************************************

Set tmWB = ActiveWorkbook
suppNMBR = FUN_suppNmbr

For Each Wb In Workbooks
    If InStr(LCase(Wb.Name), "dirt asf") Then
        Set asfwb = Wb
        Exit For
    End If
Next
        
'Rollup
'====================================================================================================
    asfwb.Activate
    Call RollemUp  '>>>>>>>>>>
    On Error GoTo errhndlNoRollups
    Range("AH:AH").SpecialCells(xlCellTypeConstants, 1).EntireRow.Select
    On Error GoTo 0
    Selection.Interior.ColorIndex = 3
    asfwb.Save '(ZeusPATH & "ASF Rollups(" & pscVar & ").xlsx")
    Selection.Delete Shift:=xlUp
noRollups:

'Import to TM
'====================================================================================================
    LastRow = Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    tmWB.Activate
    Sheets("Line Item Data").Select
    
    Range(Range("A4"), Range("A" & 2 + LastRow)).EntireRow.Insert Shift:=xlDown
    asfwb.Activate
    Range(Range("A2"), Range("O" & LastRow)).Copy
    tmWB.Activate
    Range("A4").PasteSpecial xlPasteValues
    Range("L:Q").HorizontalAlignment = xlCenter
    
    Range(Range("O4"), Range("O" & LastRow + 2)).Copy   '(duplicate UOMs)
    Range("P4").PasteSpecial xlPasteValues
    
    asfwb.Activate
    Set itemRNG = Range(Range("P2"), Range("S" & LastRow))
    itemRNG.Copy
    tmWB.Activate
    Range("Q4").PasteSpecial xlPasteValues
    Range("R:R").NumberFormat = "$#,##0"
    
    asfwb.Activate
    Set itemRNG = Range(Range("X2"), Range("X" & LastRow))
    itemRNG.Copy
    tmWB.Activate
    Range("Y4").PasteSpecial xlPasteValues
    Range("Y:Y").NumberFormat = "$#,##0"
    
    asfwb.Activate
    Set itemRNG = Range(Range("Y2"), Range("AG" & LastRow))
    itemRNG.Copy
    tmWB.Activate
    Range("AB4").PasteSpecial xlPasteValues
    
    Columns("N:N").Select
    Selection.TextToColumns Destination:=Range("N1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "General"

'Autofill HCO formulas down (but don't calculate)
'====================================================================================================
          
    'Black Lines
    Range("AL3").AutoFill Destination:=Range("AL3:AL" & LastRow + 2)
    Range("HJ3").AutoFill Destination:=Range("HJ3:HJ" & LastRow + 2)
    Range("HQ3:HS3").AutoFill Destination:=Range("HQ3:HS" & LastRow + 2)
    Range("NW3").AutoFill Destination:=Range("NW3:NW" & LastRow + 2)
    
    'mbr data
    Range("U3:AA3").AutoFill Destination:=Range("U3:AA" & LastRow + 2) '(autofill formulas)
    
    'plvling
    Range("HK3:HP3").AutoFill Destination:=Range("HK3:HP" & LastRow + 2)
    
    'Bench
    Range("NX3:OP3").AutoFill Destination:=Range("NX3:OP" & LastRow + 2)
    
    'All HCO bench
    Range(Range("OQ3"), Range("OQ3").Offset(0, suppNMBR - 1)).AutoFill Destination:=Range(Range("OQ3"), Range("OQ3").Offset(LastRow - 1, suppNMBR - 1))
    Range("OY3:OZ3").AutoFill Destination:=Range(Range("OY3"), Range("OY3").Offset(LastRow - 1, 1))
    
    'all hco red
    Range(Range("HT3"), Range("HT3").Offset(0, suppNMBR * 16)).AutoFill Destination:=Range(Range("HT3"), Range("HT3").Offset(LastRow - 1, suppNMBR * 16))
    
    'all hco yellow
    Range(Range("AM3"), Range("AM3").Offset(0, suppNMBR * 18)).AutoFill Destination:=Range(Range("AM3"), Range("AM3").Offset(LastRow - 1, suppNMBR * 18))
       
     
Range(Range("N4"), Range("N3").Offset(LastRow + 2, 0)).Interior.Color = 65535
Range(Range("A4"), Range("A3").Offset(LastRow + 2, 0)).EntireRow.Calculate
Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoRollups:
On Error GoTo 0
Resume noRollups

End Sub
Sub consolidatePricePrograms()

'sort first by col A
'-----------------------------
Call FUN_Sort(ActiveSheet.Name, Range("A2:ZA100000"), Range("A2:A100000"), 1, Range("F2:F100000"), 1)

'find tier nmbr
'-----------------------------
TierNmbr = 0
For Each c In Range(Range("K1"), Range("K1").End(xlToRight))
    If InStr(LCase(c.Value), "tier 0") > 0 Then
        TierNmbr = TierNmbr + 1
    End If
Next

LastRow = FUN_lastrow("A")

'consolidate
'-----------------------------
Range("A2").Select
consolCNT = 0
Do
consolCNT = consolCNT + 1

    Set FirstCat = ActiveCell
    mtchcats = Application.CountIf(Range("A:A"), FirstCat.Value)
        If mtchcats > 1 Then
            For i = 1 To mtchcats - 1
                Range(FirstCat.Offset(i, 10), FirstCat.Offset(i, 10 + TierNmbr - 1)).Copy
                FirstCat.Offset(0, 10).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
                Range(FirstCat.Offset(i, 0), FirstCat.Offset(i, 10 + TierNmbr - 1)).ClearContents
            Next
        End If
    FirstCat.Offset(mtchcats, 0).Select
    
Loop Until consolCNT = LastRow Or IsEmpty(ActiveCell)



End Sub


