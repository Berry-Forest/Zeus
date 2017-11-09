Attribute VB_Name = "F__Rubiks"
'[TBD]
'-------------------------
'put allcats in an array by Rval so don't have to rely on sorting sort
'set upper bound, for those that are x'd out if above upper thrshold then highlight for visual inspection, but I mean that's only if I get this thing running so that the x'd out under the bound don't have to be checked too

'common variables
'-------------------------
Public FirstCat As Range
Public AllCats As Range
Public mtchcats As Integer
Public SuppQty As Integer
Public SuppDesc As String
Public SuppMinVal As Double
Public FirstCon As Range
Public ConCol As Integer
Dim Xarray() As Integer
Dim plvlOff As Integer
Dim suppOff As Integer
Dim BenchOff As Integer
Dim tenOff As Integer
Dim NotesOffFromN As Integer

'zeus variables
'-------------------------
Dim presortFLG As Integer
Dim crawlerFLG As Integer
Dim SpecRngFLG As Integer
Dim SpecRng As Range
Dim PFstdFLG As Integer

'report variables
'-------------------------
Public plvlVarRange As Double
Public suppVarRange As Double
Public BnchVarRange As Double

'Main logic variables
'-------------------------
Dim crawledAlready As Integer
Dim multipliedAlready As Integer
Dim contextAlready As Integer
Dim SecondCNT As Integer
Dim SecondContext As Integer
Dim UOMdescCNT As Integer
Dim ttlCurrVar As Integer
Dim ttlPrevVar As Integer
Dim ttlPreVar As Integer
Dim plvlCNT As Integer
Dim SubAtsFlg As Integer
Dim ContextUOM As Integer
Dim ContextRval As Variant
Dim ContextMfg As String
Dim ContextFndflg As Integer
Dim UOMfndFLG As Integer
Dim moreContext As Integer
Dim ColBkmrk As Integer
Dim TestEAflg As Integer
Dim partialContext As Integer
Dim mincol As Integer
Dim NoResoCol As Integer
Dim eaUOMflg As Integer
Dim EAbool As Integer
Dim LastTestCol As Integer
Dim ResoCNT As Integer
Dim xrefCNT As Integer
Dim BenchCNT As Integer
Dim xrefArray() As String
Dim CrossArray() As Range
Dim BenchArray() As String
Dim varTestStrt As Range
Dim varTestRng As Range
Dim frstUOM As Range
Dim lastUOM As Range
Dim UOMset As Range
Public conpos As Integer


Sub StrtEinstein()
    
    plvlVarRange = Val(Format(ZeusForm.plvlRngSet.Value, "0.00"))
    suppVarRange = Val(Format(ZeusForm.SuppRngSet.Value, "0.00"))
    BnchVarRange = Val(Format(ZeusForm.bnchRngSet.Value, "0.00"))

    If ZeusForm.CrawlerChkBox = True Then
        crawlerFLG = 1
    Else
        crawlerFLG = 0
    End If
    
    If ZeusForm.ContextChkBox = True Then
        contextFLG = 1
    Else
        contextFLG = 0
    End If
    
    If ZeusForm.PFchkBox = True Then
        PFstdFLG = 1
    Else
        PFstdFLG = 0
    End If
    
    If Not CreateReport = True Then Call FUN_CalcOff
    Application.ScreenUpdating = False
    'Call SetApplicationStartEvents("all")   '>>>>>>>>>>
    
    If HansFLG = 1 Then Exit Sub
    '******************************
       
    If Not CreateReport = True Then
        rdybox = MsgBox("Have you saved your report?", vbYesNo)
        If rdybox = vbNo Then Exit Sub
    End If

    If ZeusForm.SpecRngBox.Visible = True Then
        SpecRngFLG = 1
        Set SpecRng = Range(SpecRngBox.Text)
    Else
        SpecRngFLG = 0
    End If
    
    If ZeusForm.SortedChkBox = True Then
        presortFLG = 1
    Else
        presortFLG = 0
    End If
    
    '[xx]Sort for Loop (I would rather have it sorted cause it's better to have it sorted this way for purposes of visual inspection anyway.  So sort stays regardless of programatic functionality.)
    '====================================================================================
    If Not presortFLG = 1 Then
        Thinking1.Show (False)
        Call FUN_Sort("Line Item Data", Range("A3:ZA100000"), Range("N3:N100000"), 1, Range("R3:R100000"), 1)
        Thinking1.Hide
    End If
    
    'Count number of preexisting variances
    '-------------------------------------------
    If SpecRngFLG = 1 Then
        SpecRng.Select
    Else
        Range(Range("N2"), Range("N2").End(xlDown)).Select
    End If
    
    ttlPreVar = WorksheetFunction.CountIf(Selection.Offset(0, plvlOff), ">" & plvlVarRange) + WorksheetFunction.CountIf(Selection.Offset(0, plvlOff), "<" & -plvlVarRange)
    For i = 0 To suppNMBR
        ttlPreVar = ttlPreVar + WorksheetFunction.CountIf(Selection.Offset(0, suppOff + (i * 18)), ">" & suppVarRange) + WorksheetFunction.CountIf(Selection.Offset(0, suppOff + (i * 18)), "<" & -suppVarRange)
        ttlPreVar = ttlPreVar + WorksheetFunction.CountIf(Selection.Offset(0, BenchOff + (i * 16)), ">" & BnchVarRange) + WorksheetFunction.CountIf(Selection.Offset(0, BenchOff + (i * 16)), "<" & -BnchVarRange)
    Next
    ttlPreVar = ttlPreVar + WorksheetFunction.CountIf(Selection.Offset(0, tenOff), ">" & BnchVarRange) + WorksheetFunction.CountIf(Selection.Offset(0, tenOff), "<" & -BnchVarRange)
    
    ttlPreVar = ttlPreVar + WorksheetFunction.CountIf(Selection.Offset(0, 4), 0) + WorksheetFunction.CountIf(Selection.Offset(0, 7), 0)
    
    For Each c In Selection.Offset(0, 3)
        If c.Offset(0, -1).Value = 1 Then
            If c.Value = "BN" Or c.Value = "CA" Or c.Value = "PL" Or c.Value = "BG" Or c.Value = "BX" Or c.Value = "CT" Or c.Value = "DZ" Or c.Value = "PK" Then ttlPreVar = ttlPreVar + 1
        End If
    Next
        
    HansEAval = 0
    
    Call CommonSetup  '>>>>>>>>>>
    Call MainLoop  '>>>>>>>>>>
    
        
End Sub
Sub CommonSetup()


'Set variables
'====================================================================================
    
    'Set common variables
    '------------------------
    Sheets("Line Item Data").Select
    suppNMBR = FUN_suppNmbr
        
    'Find column offsets
    '------------------------
    plvlOff = Range("HN2").Column - Range("N2").Column                                          '(HN offset from N)
    suppOff = Range("AM2").Column - Range("N2").Column
    BenchOff = Range("HT2").Column - Range("N2").Column
    tenOff = Range("NX2").Column - Range("N2").Column
    CalcOffFromN = Range("ZA2").Column - Range("N2").Column                                 '(ZA offset from N)

    'Set up Rubiks notes column
    '------------------------
    On Error GoTo errhndlNEWNOTES
    NotesOffFromN = Rows("2:2").Find(what:="Rubiks Notes").Column - Range("N2").Column
    On Error GoTo 0
    
    'Set up Line item research section on notes tab
    '------------------------
    If Trim(Sheets("notes").Range("K1").Value) = "" Then
        Sheets("notes").Range("K1").Value = "Line Item Research"
        Sheets("notes").Range("K1").Font.Bold = True
        Sheets("notes").Range("K1").Font.Underline = True
    End If
    
    If Not HansFLG = 1 Then
'        'DoEvents
'        'Thinking1.Show (False)
'        'DoEvents
'        HansEAval = 0
        Range(Range("N1").Offset(0, NotesOffFromN), Range("N1").Offset(FUN_lastrow("A"), NotesOffFromN).End(xlToRight)).Clear
'        Range("N2").Offset(0, NotesOffFromN).Value = "Rubiks Notes"
'        Range("N2").Offset(0, NotesOffFromN).Interior.color = 16711935
    End If
    
    'make sure formulas in col HN are setup correctly
    '====================================================================================
    If Not InStr(Range("HN2").Offset(1, 0).Formula, "LOWER") > 0 Then
        Range("HN3").Formula = "=IF(ISERROR(+HP3/X3),IF(AND(OR(R3=0,U3=0),LOWER(W3)=LOWER(""x"")),0,1),+HP3/X3)"
        Range("HN3").AutoFill Destination:=Range(Range("HN3"), Range("HN3").End(xlDown))
        Range(Range("HN3"), Range("HN3").End(xlDown)).Calculate
    End If

    Range("YX:YX").ClearContents
    Range("YX1:YX2").Value = "x"
    Range("YZ1:YZ2").Value = "x"
    plvlCNT = 0

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNEWNOTES:
NotesOffFromN = Rows("2:2").Find(what:="Spend Crossed").End(xlToRight).Offset(0, 1).Column - Range("N2").Column
Resume Next


End Sub
Sub MainLoop()
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'START LOOP //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
DoEvents
Thinking1.Show (False)
DoEvents

If SpecRngFLG = 1 Then
    SpecRng.Select
    EndRw = SpecRng.Row + SpecRng.Rows.Count - 1
Else
    Range("N3").Select
    EndRw = FUN_lastrow("N")
End If
If ActiveCell.Value = 0 Then ActiveCell.Offset(Application.CountIf(Range("N:N"), 0), 0).Select

ttlCurrVar = 0

FindNxt:
'================================================================================
If SecondContext = 1 Then
    SecondCNT = SecondCNT + 1
    If Not Range("YX2").Offset(SecondCNT, 0).Value = "" Then
        Call secondLoop
        Call MainLogic
        GoTo FindNxt
    End If
Else
    If ActiveCell.Row > EndRw Then 'Trim(ActiveCell.Value) = "" Or ActiveCell.Value = 0 Or (SpecRngFLG = 1 And ActiveCell.Row > endRw) Then
        If Not Trim(Range("YX3").Value) = "" Then
            UOMdescCNT = Application.CountA(Range(Range("YX3"), Range("YX1").End(xlDown)))  '<--Just for the status bar
            SecondContext = 1
            GoTo FindNxt
        End If
    Else
        Call VarCheck
        If Not ttlCurrVar = ttlPrevVar Then
            Call MainLogic
            FirstCat.Offset(mtchcats, 0).Select
            GoTo FindNxt
        End If
    End If
End If

Call EndClean


End Sub
Sub VarCheck()

'Define Starting Variables
'---------------------------
ttlPrevVar = ttlCurrVar
Set FirstCat = ActiveCell
mtchcats = Application.CountIf(Range("N:N"), FirstCat.Value)
ReDim Xarray(1 To mtchcats)
Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))

'check variances
'--------------------------
ttlCurrVar = ttlCurrVar + WorksheetFunction.CountIf(AllCats.Offset(0, plvlOff), ">" & plvlVarRange) + WorksheetFunction.CountIf(AllCats.Offset(0, plvlOff), "<" & -plvlVarRange)
For i = 0 To suppNMBR
    ttlCurrVar = ttlCurrVar + WorksheetFunction.CountIf(AllCats.Offset(0, suppOff + (i * 18)), ">" & suppVarRange) + WorksheetFunction.CountIf(AllCats.Offset(0, suppOff + (i * 18)), "<" & -suppVarRange)
    ttlCurrVar = ttlCurrVar + WorksheetFunction.CountIf(AllCats.Offset(0, BenchOff + (i * 16)), ">" & BnchVarRange) + WorksheetFunction.CountIf(AllCats.Offset(0, BenchOff + (i * 16)), "<" & -BnchVarRange)
Next
ttlCurrVar = ttlCurrVar + WorksheetFunction.CountIf(AllCats.Offset(0, tenOff), ">" & BnchVarRange) + WorksheetFunction.CountIf(AllCats.Offset(0, tenOff), "<" & -BnchVarRange)

'Check for 0 in Ucost and usage
'--------------------------
ttlCurrVar = ttlCurrVar + WorksheetFunction.CountIf(AllCats.Offset(0, 4), 0) + WorksheetFunction.CountIf(AllCats.Offset(0, 7), 0)

'check for pkg mismatches
'--------------------------
For Each c In AllCats.Offset(0, 3)
    If c.Offset(0, -1).Value = 1 And (c.Value = "BN" Or c.Value = "CA" Or c.Value = "PL" Or c.Value = "BG" Or c.Value = "BX" Or c.Value = "CT" Or c.Value = "DZ" Or c.Value = "PK") Then ttlCurrVar = ttlCurrVar + 1
Next

If Not HansFLG = 1 Then Application.StatusBar = "Working variances: " & ttlCurrVar & " of " & ttlPreVar & ": " & Format(ttlCurrVar / ttlPreVar, "0%")

     
End Sub
Sub secondLoop()

    'Work through seconds
    '-------------------------------------------
    Set FirstCat = Range("N:N").Find(what:=Range("YX2").Offset(SecondCNT, 0), lookat:=xlWhole)
    mtchcats = Application.CountIf(Range("N:N"), FirstCat.Value)
    ReDim Xarray(1 To mtchcats)
    Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))
    DoEvents
    Application.StatusBar = "Double Checking Pkg Strings: " & SecondCNT & " of " & UOMdescCNT
    
    
End Sub
Sub MainLogic()
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MAIN LOGIC //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    'Setup
    '-------------------------
    SuppQty = 0
    SuppMinVal = 0
    crawledAlready = 0
    multipliedAlready = 1
    contextAlready = 0
    ContextFndflg = 0
    UOMfndFLG = 0
    moreContext = 0
    ColBkmrk = 0
    TestEAflg = 0
    partialContext = 0
    NoResoCol = 0
    eaUOMflg = 0
    SubAtsFlg = 0
    
    Range("YY:BAA").ClearContents
    Range("ZA1:ZA2").Value = "x"
    ConCol = FUN_suppChk(FirstCat.Offset(0, -2).Value, FirstCat.Row)    '<--Contracted Supplier
    'TenCol = Range("NX" & FirstCat.Row).Value
    If Not HansEAval = 0 Then SuppMinVal = HansEAval
    
    'Logic Map
    '-------------------------
'    If mtchcats = 1 Then
'        Call OneCats
'        If SubAtsFlg = 1 Then GoTo SubAts
'        Exit Sub
'    Else
        
        'if no Rvals then exit
        '-------------------------------------
        Call parseRgroups
        If Trim(Range("ZA3").Value) = "" Then Exit Sub
        
        'If mbrUOMs are 0 and no other UOMs then x out all
        '-------------------------------------
        Call SetupDataGroups
        If Trim(frstUOM.Value) = "" Then
            '[TBD] check crawler/context
            AllCats.Offset(0, 9).Value = "x"
            Exit Sub
        End If
        
SubAts: Call SubsequentAttempts
        Call TestScenarios
TAgain: Call TestEAcol
        
        'if already exhausted all resources and there's no EA vals but there's a 1 then go back and check and see if that works
        '===================================================================================================================================================
        If Not ResoCNT = varTestRng.Count Then
            If EAbool = 2 And multipliedAlready = 1 And crawledAlready = 1 And contextAlready = 1 And Not TestEAflg = 1 Then
                EAbool = 1
                TestEAflg = 1
                GoTo TAgain
            Else
                mincol = 0
            End If
        End If
        
        If mincol = 0 And moreContext = 1 And crawledAlready = 1 And Not SecondContext = 1 And Not HansFLG = 1 Then
            Range("YX1").End(xlDown).Offset(1, 0).Value = FirstCat.Value
            Exit Sub
        End If
        
        'Insert New UOMs or gather more data and try again
        '========================================================================================================================================
        If Not mincol = 0 Then
            Call InsertResults
            Exit Sub
        ElseIf multipliedAlready = 0 Then
            Call MultiplyUOMs
            GoTo SubAts
        ElseIf contextAlready = 0 Then
            Call ContextUOMs
            If ContextFndflg = 1 And contextFLG = 1 Then
                moreContext = 0
                GoTo SubAts
            End If
        End If

        If crawledAlready = 0 And crawlerFLG = 1 Then
            Call CrawlThatB
            
            'cross with contextUOMs and add to UOMset if good refs
            '---------------------------
            If UOMfndFLG = 1 Then
                For Each c In Range(lastUOM.Offset(1, 0), frstUOM.Offset(-2, 0).End(xlDown))
                    If Not (Application.CountIf(AllCats.Offset(0, 1), c.Value) > 0 Or Application.CountIf(ReturnSet, c.Value) > 0) Then
                        c.ClearContents
                    End If
                Next
                If Application.CountA(Range(lastUOM.Offset(1, 0), lastUOM.End(xlDown))) > 0 Then
                    moreContext = 0
                    GoTo SubAts
                End If
            End If
        End If
        
        Call InsertResults
    
'    End If
        
        
End Sub
Sub parseRgroups()

    'separate Rvals based on price proximity and x out true 0s
    '************************************************************
            
    prevRchk = 0
    nextRchk = 0
    ColBkmrk = 0
    If Not FirstCat.Offset(0, 4).Value <= 0 Then
        Range("ZA3").Value = FirstCat.Offset(0, 4).Value
        prevRchk = 1
        If Not FirstCat.Offset(0, 6).Value > 0 Then
            FirstCat.Offset(0, 9).Value = "x"
            FirstCat.Offset(0, 9).Interior.ColorIndex = 3
        End If
    Else
        FirstCat.Offset(0, 9).Value = "x"
        FirstCat.Offset(0, 9).Interior.ColorIndex = 3
    End If

    If mtchcats = 1 Then Exit Sub
    For Each c In Range(FirstCat.Offset(1, 4), FirstCat.Offset(mtchcats - 1, 4))
        
        'Get Rval
        '------------------------------
        If c.Value > 0 Then
        
            'check col T for 0 or neg
            '------------------------------
            If Not c.Offset(0, 1).Value > 0 Then
                c.Offset(0, 5).Value = "x"
                c.Offset(0, 5).Interior.ColorIndex = 3
            End If
            
            nextRchk = 1
            Rval = c.Value
        Else
            c.Offset(0, 5).Value = "x"
            c.Offset(0, 5).Interior.ColorIndex = 3
            GoTo nxtRval
        End If
        
        'Get prev Rval (if prev is 0 then skip variance Eval, add to column, and goto next)
        '------------------------------
        If prevRchk = 1 Then
            rvalprev = c.Offset(-1, 0).Value
        Else
            If nextRchk = 1 Then prevRchk = 1   '(if the prev was 0 but curr is not then turn prev back on)
            Range("ZA1").Offset(0, ColBkmrk).End(xlDown).Offset(1, 0).Value = c.Value   '(add Rval to R group columns)
            GoTo nxtRval
        End If
        
        'separate R vals into groups based on price proximity (If the R val is under 20 and greater than threshold of the previous R value or it is more than 20 and greater than threshold+30% of the previous R value then calculate a new value)
        '------------------------------
        VarVal = (Rval - rvalprev) / Rval
        If VarVal < plvlVarRange And VarVal > -1 * plvlVarRange Then
            Range("ZA1").Offset(0, ColBkmrk).End(xlDown).Offset(1, 0).Value = c.Value
        Else
            ColBkmrk = ColBkmrk + 1
            Range("ZA1:ZA2").Offset(0, ColBkmrk).Value = "x"
            Range("ZA1").Offset(2, ColBkmrk).Value = c.Value
        End If
nxtRval:
    Next
    
End Sub
Sub SetupDataGroups()
        
    'Get avg for each R group and transpose them into single column
    '====================================================================================================================================================
    Range("ZA1:ZA2").Offset(0, ColBkmrk + 1).Value = "x"
    For i = 0 To ColBkmrk
        Range("ZA1").Offset(0, ColBkmrk + 1).End(xlDown).Offset(1, 0).Value = Application.Average(Range(Range("ZA3").Offset(0, i), Range("ZA1").Offset(0, i).End(xlDown))) 'median(Range(Range("ZA3").Offset(0, i), Range("ZA1").Offset(0, i).End(xlDown)))
    Next
    Set varTestStrt = Range("ZA3").Offset(0, ColBkmrk + 1)
    Set varTestRng = Range(varTestStrt, Range("ZA1").Offset(0, ColBkmrk + 1).End(xlDown))
    
    'Get UOMs
    '====================================================================================================================================================
    
    'Mbr
    Set frstUOM = Range("ZA2").End(xlToRight).Offset(1, 1)
    Range(frstUOM.Offset(-2, 0), frstUOM.Offset(-1, 0)).Value = "x"
    'Range(Range("ZA1").End(xlToRight), Range("ZA2").End(xlToRight)).Offset(0, 1).Value = "x"
    For Each c In AllCats.Offset(0, 1)
        If c.Value > 0 And Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), c.Value) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = c.Value
    Next
    
    'contracted supplier
    If SuppQty > 0 Then
        If Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), SuppQty) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = SuppQty
    End If
    
    'xrefs & Benchmarks
    Call CheckXrefs
    Call CheckBenches
    
    '10%
    If Not Trim(Range("NX" & FirstCat.Row).Value) = "" Then
        TenQty = Sheets("Best Market Price").Range("A:A").Find(what:=FirstCat.Value, lookat:=xlWhole).Offset(0, 10).Value
        If Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), TenQty) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = TenQty
    End If
    
    'Hans input
    If HansUOMflg = 1 Then
        For Each c In HansSet
            If Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), c.Value) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = c.Value
        Next
    End If
    
    'if there's an EA desc and no 1 already then add a 1
    '--------------------------------------
    If Application.CountIf(Range(FirstCat.Offset(0, 3), FirstCat.Offset(Application.CountA(Range(Range("ZA3"), Range("ZA1").End(xlDown))) - 1, 3)), "EA") > 0 Then
        If Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), 1) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = 1
    End If
    

End Sub
Sub SubsequentAttempts() '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    'Redefine UOMset for subsequent resolution attempts
    '****************************************************
    
    Set lastUOM = Range(Left(frstUOM.Address(0, 0), 2) & FUN_lastrow(Left(frstUOM.Address(0, 0), 2)))
    Set UOMset = Range(frstUOM, lastUOM)
    UOMset.NumberFormat = "General"
    
    'Sort UOMs ascending
    '------------------------------------
    On Error GoTo errhndlNoMoreToSort
    For i = 1 To UOMset.Count - 1
        sortminval = WorksheetFunction.Min(Range(UOMset(i), lastUOM))
        Range(UOMset(i), lastUOM).Find(what:=WorksheetFunction.Min(Range(UOMset(i), lastUOM)), lookat:=xlWhole).Value = UOMset(i).Value
        UOMset(i).Value = sortminval
    Next
NoMoreToSort:
    Set lastUOM = Range(Left(frstUOM.Address(0, 0), 2) & FUN_lastrow(Left(frstUOM.Address(0, 0), 2)))
    Set UOMset = Range(frstUOM, lastUOM)
    
    
Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNoMoreToSort:
On Error GoTo 0
Resume NoMoreToSort
    
End Sub
Sub TestScenarios()

    'Test each UOM scenario to see if any resolve
    '=======================================================================================================
    
    'If no Rval test var within Rvals
    '------------------------------------
    Range(frstUOM.Offset(-1, 1), lastUOM.End(xlToRight)).ClearContents
    If Not SuppMinVal > 0 Then
        If Not ColBkmrk = 0 Then
            For i = 1 To UOMset.Count - ColBkmrk
                Set strtuom = frstUOM.Offset(i - 1, 0)
                frstUOM.Offset(-1, i * 2).Value = strtuom.Value
                UOMcnt = Range(strtuom, lastUOM.Offset(-1, 0)).Count
                    
                'get variance between each UOM
                '------------------------------------
                For Each c In Range(varTestStrt, varTestStrt.Offset(-2, 0).End(xlDown).Offset(-1, 0))
                    c.Offset(0, i * 2).Value = 100000
                    For e = 1 To UOMcnt
                        If strtuom.Offset(e, 0).Value >= c.Offset(1, i * 2 + 1).End(xlUp).Value Then
                            EAval = c.Offset(1, 0).Value / strtuom.Offset(e, 0).Value
                            prevEAval = c.Value / c.Offset(1, i * 2 + 1).End(xlUp).Value
                            VarVal = FUN_VarVal(EAval, prevEAval)
                            If VarVal < plvlVarRange And VarVal > -plvlVarRange And Abs(VarVal) < c.Offset(0, i * 2).Value Then
                                c.Offset(0, i * 2).Value = Abs(VarVal)
                                c.Offset(0, i * 2 + 1).Value = strtuom.Offset(e, 0).Value
                            End If
                        End If
                    Next
                Next
            Next
        Else
            For i = 1 To UOMset.Count
                frstUOM.Offset(-1, i * 2).Value = frstUOM.Offset(i - 1, 0).Value
                frstUOM.Offset(0, i * 2 - 1).Value = 0.01
            Next
        End If
    Else
        
        'If supplier minval test var from suppminval
        '------------------------------------
        For Each c In varTestRng
            c.Offset(0, 2).Value = 100000
            For Each d In UOMset
                EAval = c.Value / d.Value
                VarVal = FUN_VarVal(EAval, SuppMinVal)
                If VarVal < suppVarRange And VarVal > -suppVarRange And Abs(VarVal) < c.Offset(0, 2).Value Then
                    c.Offset(0, 2).Value = Abs(VarVal)
                    c.Offset(-1, 3).Value = d.Value
                End If
            Next
        Next
    End If
    
    'if there's a one but no EA desc then try to use any other group
    '-------------------------------
    If Application.CountIf(AllCats.Offset(0, 3), "EA") = 0 And Application.CountIf(UOMset, 1) > 0 And Not SuppMinVal > 0 Then
        EAbool = 2
    Else
        EAbool = 1
    End If
    
    If SuppMinVal > 0 Then
        LastTestCol = frstUOM.Offset(0, 2).Column
    Else
        LastTestCol = frstUOM.Offset(0, (UOMset.Count - ColBkmrk) * 2).Column
    End If
    'lasttestcol = FUN_lastcol(frstUOM.Offset(-1, 0).Row)

End Sub
Sub TestEAcol()

    'Test to see which scenario resolves closest
    '===================================================================================================================================================
    If 100000 * (varTestRng.Count - 1) > 0 Then
        minabs = 100000 * (varTestRng.Count - 1)
    Else
        minabs = 100000
    End If
    mincol = 0
    ResoCNT = 0
    For i = EAbool To (LastTestCol - frstUOM.Column) / 2
        Set resorng = Range(frstUOM.Offset(-1, i * 2), frstUOM.Offset(varTestRng.Count - 1, i * 2))
        If Application.CountA(resorng) >= ResoCNT Then
            ResoCNT = Application.CountA(resorng)
            
            For Each c In resorng.Offset(0, -1)
                If c.Value = 100000 Then c.ClearContents
            Next
            If Not Application.CountA(resorng.Offset(0, -1)) = 0 Then
            
                'if total var is less than prevtotalvar and is significantly distanced
                '------------------------------------
                totalvar = Application.sum(resorng.Offset(0, -1))
                If totalvar = 0 Then
                    If totalvar = 0 And minabs = 0 Then
                        GoTo PrevCurrChk
                    Else
                        minabs = totalvar
                        NoResoCol = i
                        mincol = i
                        GoTo nxtTstCol
                    End If
                End If
                If totalvar < minabs And ((totalvar - minabs) / totalvar > 0.25 Or (totalvar - minabs) / totalvar < -0.25) Then
                    minabs = totalvar
                    NoResoCol = i
                    mincol = i
                ElseIf ((totalvar - minabs) / totalvar < 0.25 Or (totalvar - minabs) / totalvar > -0.25) Then
PrevCurrChk:
                    'if any UOMs match contextUOM then use that
                    '------------------------------------
                    If contextAlready = 0 And partialContext = 0 Then
                        Range("YY1:YZ2").Value = "x"
                        ContextMfg = FirstCat.Offset(0, -2).Value
                        Set ReturnStrt = Range("YY3")
                        For Each c In varTestRng
                            ContextRval = c.Value
                            Call ContextCheck
                        Next
                        partialContext = 1
                    End If
                    
                    'find most frequent conUOMs that match present UOMs
                    '------------------------------------
                    prevmintest = 0
                    currmintest = 0
                    For conCompare = 1 To varTestRng.Count
                        On Error Resume Next
                        prevmintest = prevmintest + ReturnSet.Find(what:=frstUOM.Offset(conCompare - 2, mincol * 2).Value, lookat:=xlWhole).Offset(0, 1).Value
                        currmintest = currmintest + ReturnSet.Find(what:=frstUOM.Offset(conCompare - 2, i * 2).Value, lookat:=xlWhole).Offset(0, 1).Value
                        On Error GoTo 0
                    Next

                    If Not currmintest = 0 Then
                        If (currmintest - prevmintest) / currmintest > 0.25 Then
                            minabs = totalvar
                            NoResoCol = i
                            mincol = i
                        ElseIf SecondContext = 1 Or HansFLG = 1 Then
                            If currmintest > prevmintest Then
                                minabs = totalvar
                                NoResoCol = i
                            Else
                                If Application.CountIf(AllCats.Offset(0, 3), "EA") + Application.CountIf(AllCats.Offset(0, 3), "RL") < mtchcats * 0.5 Then
                                    minabs = totalvar
                                    NoResoCol = i
                                End If
                                moreContext = 1
                                mincol = 0
                            End If

                        ElseIf (currmintest - prevmintest) / currmintest > -0.25 Then
                            moreContext = 1
                            mincol = 0
                        End If
                    ElseIf prevmintest = 0 Then
                        If totalvar < minabs Then
                            minabs = totalvar
                            NoResoCol = i
                        End If
                        moreContext = 1
                        mincol = 0
                    End If
                    
                End If
            End If
        End If
nxtTstCol:
    Next

End Sub
Sub MultiplyUOMs()

'
'        'multiply UOMs to get hidden qtys
'        '------------------------------------
'        '[?] maybe see if multiplier that resolve do so within certain threshold, otherwise use next closest and x out
'        For Each c In UOMset
'            For Each D In UOMset
'                If Not Application.CountIf(Range(frstUOM, frstUOM.Offset(-2, 0).End(xlDown)), c.Value * D.Value) > 0 Then frstUOM.Offset(-2, 0).End(xlDown).Offset(1, 0).Value = c.Value * D.Value
'            Next
'        Next
'        multipliedalready = 1


End Sub
Sub ContextUOMs()

    'Find Context UOMs
    '===================================================================================================================================================
    
        'find context UOMs
        '------------------------------------
        Range(Range("YY3"), Range("YZ1").End(xlDown)).ClearContents
        Range("YY1:YZ2").Value = "x"
        ContextMfg = FirstCat.Offset(0, -2).Value
        For Each c In varTestRng
            Set ReturnStrt = Range("YY1").End(xlDown).Offset(1, 0)
            ContextRval = c.Value
            Call ContextCheck
'~~
            If Not ReturnStrt.Value = "" Then
                For Each d In Range(ReturnStrt, Range("YY1").End(xlDown))
                    If d.Value > 0 And d.Offset(0, 1).Value > Application.Percentile(Range(ReturnStrt.Offset(0, 1), Range("YY1").End(xlDown).Offset(0, 1)), 0.9) And Not Application.CountIf(UOMset, d.Value) > 0 Then
                        frstUOM.Offset(-2, 0).End(xlDown).Offset(1, 0).Value = d.Value
                        ContextFndflg = 1
                    End If
                Next
            End If
'~~
        Next
        Set ReturnStrt = Range("YY3")
        Set ReturnSet = Range(ReturnStrt, Range("YY1").End(xlDown))
        For Each c In ReturnSet
            If Not c.Value = "" Then
                For Each d In ReturnSet
                    If c.Value = d.Value And Not c.Address = d.Address Then
                        c.Offset(0, 1).Value = c.Offset(0, 1).Value + d.Offset(0, 1).Value
                        Range(d, d.Offset(0, 1)).ClearContents
                    End If
                Next
            End If
        Next
        contextAlready = 1
'~~

'~~

    
    

End Sub
Sub CrawlThatB()
       
        'lookup UOMs
        '---------------------------
        ProdNmbr = FirstCat.Value
        Set UOMref = FirstCat.Offset(0, NotesOffFromN)
        Range(UOMref, UOMref.End(xlToRight)).ClearContents
        UOMref.Value = "x"   'placeholder for einstein notes
        'Range("A1:A2").Offset(0, lasttestcol).Value = "x"
        'Set ReturnCol = Range("A3").Offset(0, lasttestcol)
        Set ReturnCol = frstUOM
        Call UOMcrawler_CIA  '>>>>>>>>>>
        crawledAlready = 1
        
        'input research notes on Notes tab
        '---------------------------
        If Not IsEmpty(Sheets("notes").Range("K2")) Then
            Sheets("notes").Range("K1").End(xlDown).Offset(1, 1).Value = urlLnk
            Sheets("notes").Range("K1").End(xlDown).Offset(1, 0).Value = ProdNmbr
        Else
            Sheets("notes").Range("K2").Offset(0, 1).Value = urlLnk
            Sheets("notes").Range("K2").Value = ProdNmbr
        End If
        

        

End Sub
Sub InsertResults()

    'if no resolution use NoReso as min col to insert closest UOMs and x out, well what about the fact that the test will only build ascending and stop when not found.  Maybe take UOM set that's been built and send to baseline and calculation section, no crawler.
    '===========================================================================================================
    '[?] if variance is less than 80% then x out is legit, otherwise highlight it for inspection?


    If Not NoResoCol > 0 Then
    
        'if no resolutions at all were found then keep existing UOMs and see which ones need to be x'out
        '------------------------------------
        
'        Call HandleConSupp
'        Call HandleXrefs
'        Call HandleBenchAndTen
        
'        'find lowest plvl Val
'        For Each c In AllCats.Offset(0, 4)
'            If Not c.Value = 0 Then
'                MinR = c.Value
'                Exit For
'            End If
'        Next
''        If SuppQty > 0 And HansEAval = 0 Then
''            SuppMinVal = FirstCat.Offset(0, ConCol + 6 - FirstCat.Column).Value / SuppQty
''        End If
'
'
'        For Each c In AllCats.Offset(0, 4)
'            If Not c.Value = 0 Then
'                If (c.Value - MinR) / c.Value > plvlVarRange Or (c.Value - MinR) / c.Value < -1 * plvlVarRange Then
'                    UnresolvCNT = UnresolvCNT + 1
'                    Xarray(c.Row - FirstCat.Row + 1) = 1
''                ElseIf SuppMinVal > 0 Then
''                    If ((c.Value / c.Offset(0, -2)) - SuppMinVal) / (c.Value / c.Offset(0, -2)) > suppVarRange Or ((c.Value / c.Offset(0, -2)) - SuppMinVal) / (c.Value / c.Offset(0, -2)) - 1 * suppVarRange Then
''                        UnresolvCNT = UnresolvCNT + 1
''                        Xarray(c.Row - FirstCat.Row + 1) = 1
''                    End If
'                End If
'            End If
'        Next
        
    Else
        mincol = NoResoCol
        Set frstmincol = frstUOM.Offset(-1, mincol * 2)
        
        'fill in the UOMS that were not found
        '---------------------------
        'If Not Application.CountA(frstmincol.EntireColumn) = vartestrng.Count - 1 Then
            If frstmincol.Value = "" Then
                frstmincol.Value = frstmincol.Offset(-1, 0).End(xlDown).Value
            End If
            If Not varTestRng.Count = 1 Then
                For Each c In Range(frstmincol.Offset(1, 0), frstmincol.Offset(ColBkmrk, 0))
                    If c.Value = "" Then
                        If Not c.Address = frstmincol.Offset(ColBkmrk, 0).Address Then
                            currR = c.Offset(1, (-2 * mincol) - 1).Value
                            prevR = Range("ZA1").Offset(0, c.Row - 3).End(xlDown).Value
                            nxtR = Range("ZA3").Offset(0, c.Row - 1).Value
                            prevUOM = c.Offset(-1, 0).Value
                            nxtUOM = c.Offset(1, 0).Value
                            If Abs(((currR / prevUOM) - (prevR / prevUOM)) / (currR / prevUOM)) < Abs(((currR / nxtUOM) - (nxtR / nxtUOM)) / (currR / nxtUOM)) Then
                                c.Value = c.Offset(-1, 0).Value
                            Else
                                c.Value = c.Offset(1, 0).Value
                            End If
                        Else
                            c.Value = c.Offset(-1, 0).Value
                        End If
                    End If
                Next
            End If
        'Else
            'Set frstmincol = frstmincol.Offset(-1, 0).End(xlDown)
        'End If
        
'        'get suppminval if not used to calculate already
'        '---------------------------
'        If SuppQty > 0 And HansEAval = 0 Then
'            SuppMinVal = FirstCat.Offset(0, ConCol + 6 - FirstCat.Column).Value / SuppQty
'        End If
        
        'input new values
        '---------------------------
        For Each c In AllCats.Offset(0, 4)
            If Not c.Value > 0 Then
                c.Offset(0, -2).Value = frstmincol.Value
            Else
                For i = 0 To ColBkmrk
                    If Application.CountIf(Range(Range("ZA3").Offset(0, i), Range("ZA3").Offset(-2, i).End(xlDown)), c.Value) > 0 Then
                        If Not c.Offset(0, -2).Value = frstmincol.Offset(i, 0).Value And Not frstmincol.Offset(i, 0).Value = "" Then
                            c.Offset(0, -2).Value = frstmincol.Offset(i, 0).Value
                        
                            'mark UOM method
                            '------------------------------------
                            If Not Application.CountIf(AllCats.Offset(0, 1), c.Offset(0, -2).Value) > 0 Then
                                If c.Offset(0, -2).Value = SuppQty Then
                                    c.Offset(0, 4).Interior.ColorIndex = 39
                                ElseIf UOMfndFLG = 1 Then
                                    For Each crawlU In Range(Range(UOMref.Address), Range(UOMref.Address).End(xlToRight))
                                        If InStr(crawlU.Value, c.Offset(0, -2).Value & "/") > 0 Then
                                            c.Offset(0, 4).Interior.ColorIndex = 8
                                            Exit For
                                        End If
                                    Next
                                ElseIf Application.CountIf(ReturnSet, c.Offset(0, -2).Value) > 0 Then
                                    c.Offset(0, 4).Interior.ColorIndex = 36
                                End If
                            End If
                        
                        End If
                        
'                        If SuppMinVal > 0 Then
'                            If ((c.Value / c.Offset(0, -2)) - SuppMinVal) / (c.Value / c.Offset(0, -2)) > suppVarRange Or ((c.Value / c.Offset(0, -2)) - SuppMinVal) / (c.Value / c.Offset(0, -2)) < -1 * suppVarRange Then
'                                UnresolvCNT = UnresolvCNT + 1
'                                Xarray(c.Row - FirstCat.Row + 1) = 1
'                            End If
'                        End If
                        
                        Exit For
                    End If
                Next
            End If
        Next
    
    End If
        
        Call HandleConSupp
        Call HandleXrefs
        Call HandleBenchAndTen
        
        'find lowest EAval in allcats
        '---------------------------
        minEAval = 100000
        For Each c In AllCats.Offset(0, 4)
            If Not c.Offset(0, 5).Value = "x" And c.Value / c.Offset(0, -2).Value < minEAval Then
                minEAval = c.Value / c.Offset(0, -2).Value
            End If
        Next
        
        'xout plvlvar and insert x's from Xarray
        '---------------------------
        For i = 0 To mtchcats - 1
            If Xarray(i + 1) = 1 Then
                FirstCat.Offset(i, 9).Value = "x"
            ElseIf Not FirstCat.Offset(i, 4).Value = 0 Then
                VarVal = FUN_VarVal(FirstCat.Offset(i, 4).Value / FirstCat.Offset(i, 2), minEAval)
                If VarVal > plvlVarRange Or VarVal < -plvlVarRange Then
                    UnresolvCNT = UnresolvCNT + 1
                    FirstCat.Offset(i, 9).Value = "x"
                End If
            End If
        Next


    
End Sub
Sub HandleConSupp()

'[TBD]call after new allcats vals have been inserted but before x out

If ConCol = 0 Then Exit Sub

For Each c In AllCats
    'If LCase(FirstCat.Offset(I, 9).Value) = "x" Then
        'If Range(FirstCon).Offset(I, 4).Value = 1 Then
            'If c.Value = "BN" Or c.Value = "CA" Or c.Value = "PL" Or c.Value = "BG" Or c.Value = "BX" Or c.Value = "CT" Or c.Value = "DZ" Or c.Value = "PK" Then
                SuppMinVal = FirstCon.Offset(0, 6).Value
                VarVal = FUN_VarVal(c.Offset(0, 4).Value / c.Offset(0, 2).Value, SuppMinVal)
                If VarVal > 1 Or VarVal < -1 Then
                        
                    'determine if at least half are > 100% and > 3/4 have high variances
                    '-------------------------------------
                    For Each d In AllCats
                        VarVal = FUN_VarVal(d.Offset(0, 4).Value / d.Offset(0, 2).Value, SuppMinVal)
                        If VarVal > suppVarRange Or VarVal < -suppVarRange Then unResoCNT = unResoCNT + 1
                        If VarVal > 1 Or VarVal < -1 Then HighVarCnt = HighVarCnt + 1
                    Next
                    If Not HighVarCnt / mtchcats >= 0.5 Or Not unResoCNT / mtchcats > 0.75 Then Exit Sub
                            
'                    'determine by total variance
'                    '-------------------------------------
'                    VarMin = 100000
'                    For Each UOMtestQty In UOMset
'                        ttlTestVar = 0
'                        For j = 0 To mtchcats
'                            varval = FUN_VarVal(Range(FirstCon).Offset(I, 6).Value / UOMtestQty, FirstCat.Offset(I, 4).Value / FirstCat.Offset(I, 2).Value)
'                            ttlTestVar = ttlTestVar + varval
'                        Next
'                        If ttlTestVar < VarMax Then
'                            VarMin = Abs(ttlTestVar)
'                            TempUOM = UOMtestQty
'                        End If
'                    Next
                
                    '-OR-
                    
                    'determine by how many are under the threshold first, then by total variance
                    '-------------------------------------
                    VarMin = 100000
                    ConResoMax = 0
                    For Each UOMtestQty In UOMset
                        ttlTestVar = 0
                        ConResoNmbr = 0
                        TestMinVal = FirstCon.Offset(0, 6).Value / UOMtestQty
                        For Each e In AllCats
                            VarVal = FUN_VarVal(e.Offset(0, 4).Value / e.Offset(0, 2).Value, TestMinVal)
                            ttlTestVar = ttlTestVar + VarVal
                            If VarVal < suppVarRange And VarVal > -suppVarRange Then ConResoNmbr = ConResoNmbr + 1
                        Next
                        If ConResoNmbr > ConResoMax Then
                            ConResoMax = ConResoNmbr
                            VarMin = Abs(ttlTestVar)
                            tempUOM = UOMtestQty
                        ElseIf ConResoNmbr = ConResoMax And ttlTestVar < VarMin Then
                            VarMin = Abs(ttlTestVar)
                            tempUOM = UOMtestQty
                        End If
                    Next
                    
                    'Capture which ones need to be x'd out using new UOM
                    '-------------------------------------
                    TestMinVal = FirstCon.Offset(0, 6).Value / tempUOM
                    For X = 1 To mtchcats
                        VarVal = FUN_VarVal(FirstCat.Offset(X - 1, 4).Value / FirstCat.Offset(X - 1, 2).Value, TestMinVal)
                        If VarVal > suppVarRange Or VarVal < -suppVarRange Then Xarray(X) = 1
                    Next
                    
                    'insert new values
                    '-------------------------------------
                    For Each sht In ActiveWorkbook.Sheets
                        If InStr(LCase(sht.Name), "pricing") > 0 And Not InStr(LCase(sht.Name), "pricing") = 1 Then
                            shtCNT = shtCNT + 1
                            If shtCNT = conpos Then
                                sht.Range("A:A").Find(what:=FirstCon.Offset(0, 1).Value, lookat:=xlWhole).Offset(0, 4).Value = tempUOM
                                sht.Range("A:A").Find(what:=FirstCon.Offset(0, 1).Value, lookat:=xlWhole).EntireRow.Calculate
                                SuppQty = tempUOM
                                SuppMinVal = FirstCon.Offset(0, 6).Value / SuppQty
                                Exit Sub
                            End If
                        End If
                    Next
                    
                End If
Next

End Sub
Sub HandleXrefs()

'only change in xref tab, maybe search xref col and con col for other uses of catnmbr, and if none or other instances also have variances, and no reso on xref tab then go through handlePricefile methodology to try and change
'*********************************************************************************************

For xref = 1 To xrefCNT
    Set FirstXref = Range(xrefArray(xref))
    XrefMinVal = FirstXref.Offset(0, 6).Value / FirstXref.Offset(0, 4).Value
    For Each c In AllCats
        VarVal = FUN_VarVal(c.Offset(0, 4).Value / c.Offset(0, 2).Value, XrefMinVal)
        If VarVal > suppVarRange Or VarVal < -suppVarRange Then
                        
            'find xref page and count how number of crosses
            '-------------------------------------
            mfgpos = (FirstXref.Column - Range("U1").Column) / 18
            For Each sht In ActiveWorkbook.Sheets
                If InStr(LCase(sht.Name), "cross reference") > 0 And Not InStr(LCase(sht.Name), "cross reference") = 1 Then
                    shtCNT = shtCNT + 1
                    If shtCNT = mfgpos Then
                        testNmbr = Application.CountIf(sht.Range("A:A"), FirstCat.Value)
                        ReDim CrossArray(1 To testNmbr)
                        suppAlias = Trim(Replace(LCase(sht.Name), "cross reference", ""))
                        VarMin = 100000
                        ConResoMax = 0
                        
                        'determine by how many are under the threshold first, then by total variance
                        '-------------------------------------
                        For tst = 1 To testNmbr
                            ttlTestVar = 0
                            ConResoNmbr = 0
                            Set CrossArray(tst) = sht.Range("A:A").Find(what:=FirstCat.Value, lookat:=xlWhole)
                            pfAdd = Sheets(suppAlias & " Pricing").Range("A:A").Find(what:=CrossArray(tst).Offset(0, 1).Value, lookat:=xlWhole).Address
                            TestMinVal = Sheets(suppAlias & " Pricing").Range(pfAdd).Offset(0, 9).Value / Sheets(suppAlias & " Pricing").Range(pfAdd).Offset(0, 4).Value
                            
                            For Each d In AllCats
                                VarVal = FUN_VarVal(d.Offset(0, 4).Value / d.Offset(0, 2).Value, TestMinVal)
                                ttlTestVar = ttlTestVar + VarVal
                                If VarVal < suppVarRange And VarVal > -suppVarRange Then ConResoNmbr = ConResoNmbr + 1
                            Next
                            If ConResoNmbr > ConResoMax Then
                                ConResoMax = ConResoNmbr
                                VarMin = Abs(ttlTestVar)
                                TempCross = tst
                                TempMinVal = TestMinVal
                            ElseIf ConResoNmbr = ConResoMax And ttlTestVar < VarMin Then
                                VarMin = Abs(ttlTestVar)
                                TempCross = tst
                                TempMinVal = TestMinVal
                            End If
                        Next
                        
                        'Capture which ones need to be x'd out using new UOM
                        '-------------------------------------
                        TestMinVal = TempMinVal
                        For X = 1 To mtchcats
                            VarVal = FUN_VarVal(FirstCat.Offset(X - 1, 4).Value / FirstCat.Offset(X - 1, 2).Value, TestMinVal)
                            If VarVal > suppVarRange Or VarVal < -suppVarRange Then Xarray(X) = 1
                        Next
                        
                        'Insert new values
                        '-------------------------------------
                        lastcol = sht.Range("A1").End(xlToRight).Column
                        For tst = 1 To testNmbr
                            If Not tst = TempCross Then
                                CrossArray(tst).Offset(0, lastcol).Value = CrossArray(tst).Value
                                CrossArray(tst).ClearContents
                            End If
                        Next
                        GoTo NxtXref
                        
                    End If
                End If
            Next
        End If
    Next
NxtXref:
Next


End Sub
Sub HandleBenchAndTen()

For Bench = 1 To BenchCNT + 1
    If Bench = BenchCNT + 1 Then
        If Trim(Range("NX" & FirstCat.Row).Value) = "" Then Exit Sub
        Set FirstBench = Range("NX" & FirstCat.Row)
        BMPCatNmbr = FirstCat.Value
    Else
        Set FirstBench = Range(BenchArray(Bench))
        mfgpos = (FirstBench.Column - Range("HD1").Column) / 16
        BMPCatNmbr = Range("AN" & FirstCat.Row).Offset(0, (mfgpos - 1) * 18).Value
    End If
    BenchMinVal = FirstBench.Offset(0, 1).Value
    For Each c In AllCats
        If c.Offset(0, 4).Value = 0 Then GoTo nxtC
        VarVal = FUN_VarVal(c.Offset(0, 4).Value / c.Offset(0, 2).Value, BenchMinVal)
        If VarVal > BnchVarRange Or VarVal < -BnchVarRange Then
            
            'determine by how many are under the threshold first, then by total variance
            '-------------------------------------
            VarMin = 100000
            ConResoMax = 0
            For Each UOMtestQty In UOMset
                ttlTestVar = 0
                ConResoNmbr = 0
                TestMinVal = FirstBench.Value / UOMtestQty
                For Each e In AllCats
                    If e.Offset(0, 4).Value = 0 Then GoTo nxtE
                    VarVal = FUN_VarVal(e.Offset(0, 4).Value / e.Offset(0, 2).Value, TestMinVal)
                    ttlTestVar = ttlTestVar + VarVal
                    If VarVal < BnchVarRange And VarVal > -BnchVarRange Then ConResoNmbr = ConResoNmbr + 1
nxtE:           Next
                If ConResoNmbr > ConResoMax Then
                    ConResoMax = ConResoNmbr
                    VarMin = Abs(ttlTestVar)
                    tempUOM = UOMtestQty
                ElseIf ConResoNmbr = ConResoMax And ttlTestVar < VarMin Then
                    VarMin = Abs(ttlTestVar)
                    tempUOM = UOMtestQty
                End If
            Next
            
            'Capture which ones need to be x'd out using new UOM
            '-------------------------------------
            TestMinVal = FirstBench.Value / tempUOM
            For X = 1 To mtchcats
                If FirstCat.Offset(X - 1, 4).Value = 0 Then GoTo nxtX
                VarVal = FUN_VarVal(FirstCat.Offset(X - 1, 4).Value / FirstCat.Offset(X - 1, 2).Value, TestMinVal)
                If VarVal > BnchVarRange Or VarVal < -BnchVarRange Then Xarray(X) = 1
nxtX:       Next
            
            'insert new values
            '-------------------------------------
            Sheets("Best Market Price").Range("A:A").Find(what:=BMPCatNmbr, lookat:=xlWhole).Offset(0, 10).Value = tempUOM
            Sheets("Best Market Price").Range("A:A").Find(what:=BMPCatNmbr, lookat:=xlWhole).EntireRow.Calculate
            Exit For
            
nxtC:   End If
    Next
Next

End Sub
Sub EndClean()
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'END AND CLEAN////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'=================================================================================================================================================

    'Clean up
    '-----------------------------
    Range("W:W").NumberFormat = "$#,##0"
    Range("W:W").HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Range("N3").Select
    
    If Not (CreateReport = True Or DragonSlayerFLG = 1) Then
        presortFLG = 0
    End If
    
    Unload Thinking1

    If Not (CreateReport = True Or DragonSlayerFLG = 1) Then
        MsgBox (ttlPreVar - UnresolvCNT & " of " & ttlPreVar & ": " & Format(1 - (UnresolvCNT / ttlPreVar), "0%") & " variances resolved.")
    End If
'    On Error Resume Next
'    MsgBox ("     " & pvarttl - ttlvar & "of" & ttlprevar & " variances resolved (" & Format((pvarttl - ttlvar) / ttlprevar, "0%") & ")" & vbCrLf & _
'            vbCrLf & _
'            "       " & ttlvar & " Items X'd out for validation") '& _
'            vbCrLf & _
'            "       " & suppreso & " Supplier variances unresolved" & _
'            vbCrLf & _
'            vbCrLf & _
'            "             " & pvarone + pvartwo + pvarthree & " Items for inspection:" & _
'            vbCrLf & _
'            "             " & pvarone & " 1st degree matches" & _
'            vbCrLf & _
'            "             " & pvartwo & " 2nd degree matches" & _
'            vbCrLf & _
'            "             " & pvarthree & " 3rd degree matches")
'
    If UnresolvCNT < 20 And Not (CreateReport = True Or DragonSlayerFLG = 1) Then
        Chuck.Show
    End If
    
    If Not (CreateReport = True Or DragonSlayerFLG = 1) Then
        Call FUN_CalcBackOn
    End If


End Sub
Sub OneCats()
        
        'if true 0 then x out
        '------------------------------------
        If FirstCat.Offset(0, 4).Value = 0 Or FirstCat.Offset(0, 6).Value = 0 Then
            FirstCat.Offset(0, 9).Value = "x"
            FirstCat.Offset(0, 9).Interior.ColorIndex = 3
            Exit Sub
        End If
        
        
        
        
        
        'if UOM is 0 then
        '------------------------------------
        ElseIf FirstCat.Offset(0, 1).Value = 0 Then
            
            'try to use input from Hans
            '------------------------------------
            
                
            End If
            
            'Try to pull from Con Supplier
            '------------------------------------
            If ConCol > 0 Then
                If HansEAval = 0 Then
                    If Not PFstdFLG = 1 Then
                        If SuppQty = 1 And Not (SuppDesc = "BN" Or SuppDesc = "CA" Or SuppDesc = "PL" Or SuppDesc = "BG" Or SuppDesc = "BX" Or SuppDesc = "CT" Or SuppDesc = "DZ" Or SuppDesc = "PK") Then
                            SuppMinVal = 0
                        Else
                            SuppMinVal = FirstCat.Offset(0, ConCol + 7 - FirstCat.Column).Value
                            MbrEaVal = FirstCat.Offset(0, 4).Value / SuppQty
                            If (MbrEaVal - SuppMinVal) / MbrEaVal < suppVarRange And (MbrEaVal - SuppMinVal) / MbrEaVal > -suppVarRange Then
                                FirstCat.Offset(0, 2).Value = SuppQty
                                FirstCat.Offset(0, 8).Interior.ColorIndex = 39
                                Exit Sub
                            End If
                        End If
                    Else
                        SuppMinVal = FirstCon.Offset(0, 7).Value
                        If ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) < suppVarRange And ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) > -suppVarRange Then
                            If Not FirstCat.Offset(0, 2).Value = SuppQty Then
                                FirstCat.Offset(0, 2).Value = SuppQty
                                FirstCat.Offset(0, 8).Interior.ColorIndex = 39
                            End If
                            Exit Sub
                        End If
                    End If
                Else
                    SuppMinVal = HansEAval
                    If ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) < plvlVarRange And ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) > -1 * plvlVarRange Then
                        If Not FirstCat.Offset(0, 2).Value = SuppQty Then
                            FirstCat.Offset(0, 2).Value = SuppQty
                            FirstCat.Offset(0, 8).Interior.ColorIndex = 39
                        End If
                        Exit Sub
                    End If
                End If
                Range("ZB3").Value = SuppQty
                Set frstUOM = Range("ZB3")
            Else
                '[TBD] try to pull from xrefs,benchs,Ten
                'take into account possible hansEAval?
            End If
            Range("ZA1:ZB2").Value = "x"
            Range("ZA3").Value = FirstCat.Offset(0, 4).Value
            Set varTestRng = Range("ZA3")
            Set varTestStrt = Range("ZA3")
            
            'check context
            '------------------------------------
            Range(Range("YY3"), Range("YZ1").End(xlDown)).ClearContents
            Range("YY1:YZ2").Value = "x"
            ContextMfg = FirstCat.Offset(0, -2).Value
            Set ReturnStrt = Range("YY3")
            ContextRval = FirstCat.Offset(0, 4).Value
            Call ContextCheck
            contextAlready = 1
            If Not ReturnStrt.Value = "" Then
                If SuppMinVal > 0 Then
                    For Each c In ReturnSet
                        If c.Value > 0 And Not Application.CountIf(UOMset, c.Value) > 0 Then frstUOM.Offset(-2, 0).End(xlDown).Offset(1, 0).Value = c.Value
                    Next
                    ContextFndflg = 1
                    SubAtsFlg = 1
                    Exit Sub
                Else
                    MaxFreq = WorksheetFunction.Max(ReturnSet.Offset(0, 1))
                    FirstCat.Offset(0, 2).Value = ReturnSet.Offset(0, 1).Find(what:=MaxFreq, lookat:=xlWhole, LookIn:=xlFormulas).Offset(0, -1).Value
                    Exit Sub
                End If
            End If
        
            'lookup UOMs
            '---------------------------
            ProdNmbr = FirstCat.Value
            Set UOMref = FirstCat.Offset(0, NotesOffFromN)
            Range(UOMref, UOMref.End(xlToRight)).ClearContents
            UOMref.Value = "x"   'placeholder for einstein notes
            Set ReturnCol = Range("ZB3")
            Call UOMcrawler_CIA  '>>>>>>>>>>
            crawledAlready = 1
            If UOMfndFLG = 1 Then
                Set frstUOM = Range("ZB3")
                SubAtsFlg = 1
                Exit Sub
            End If

            'if no context and no crawler
            '---------------------------
            If Not SuppQty = 0 Then
                FirstCat.Offset(0, 2).Value = SuppQty
                UnresolvCNT = UnresolvCNT + 1
                FirstCat.Offset(0, 9).Value = "x"
                Exit Sub
            Else
                FirstCat.Offset(0, 2).Value = 1
                Exit Sub
            End If

        Else
            
            'set supplier variables
            '------------------------------------
            If ConCol > 0 Then
                SuppQty = FirstCat.Offset(0, ConCol + 4 - FirstCat.Column).Value
                SuppDesc = FirstCat.Offset(0, ConCol + 5 - FirstCat.Column).Value
                If HansEAval = 0 Then
                    If Not PFstdFLG = 1 Then
                        If SuppQty = 1 And Not (SuppDesc = "EA" Or SuppDesc = "RL") Then
                            UnresolvCNT = UnresolvCNT + 1
                            FirstCat.Offset(0, 9).Value = "x"
                            '[?] maybe make a certain color to denote pricefile error
                            Exit Sub
                        Else
                            SuppMinVal = FirstCat.Offset(0, ConCol + 7 - FirstCat.Column).Value
                            If ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) < plvlVarRange And ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) > -1 * plvlVarRange Then
                                If Not FirstCat.Offset(0, 2).Value = SuppQty Then
                                    FirstCat.Offset(0, 2).Value = SuppQty
                                    FirstCat.Offset(0, 8).Interior.ColorIndex = 39
                                End If
                                Exit Sub
                            End If
                        End If
                    Else
                        SuppMinVal = FirstCat.Offset(0, ConCol + 7 - FirstCat.Column).Value
                        If ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) < plvlVarRange And ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) > -1 * plvlVarRange Then
                            If Not FirstCat.Offset(0, 2).Value = SuppQty Then
                                FirstCat.Offset(0, 2).Value = SuppQty
                                FirstCat.Offset(0, 8).Interior.ColorIndex = 39
                            End If
                            Exit Sub
                        End If
                    End If
                Else
                    SuppMinVal = HansEAval
                    If ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) < plvlVarRange And ((FirstCat.Offset(0, 4).Value / SuppQty) - SuppMinVal) / (FirstCat.Offset(0, 4).Value / SuppQty) > -1 * plvlVarRange Then
                        If Not FirstCat.Offset(0, 2).Value = SuppQty Then
                            FirstCat.Offset(0, 2).Value = SuppQty
                            FirstCat.Offset(0, 8).Interior.ColorIndex = 39
                        End If
                        Exit Sub
                    End If
                End If
            Else
                
                'if there's no supplier then there's no suppvar or plvlvar to evaluate so if there's a hans UOM then just use that automatically.
                '--------------------------------
                If HansEAval = 0 Then
                    If HansUOMflg = 0 Then
                        Exit Sub
                    Else
                        FirstCat.Offset(0, 2).Value = Range("YW3").Value
                        Exit Sub
                    End If
                Else
                    SuppMinVal = HansEAval
                    SubAtsFlg = 1
                    Exit Sub
                End If
    
            End If
            Range("ZA1:ZB2").Value = "x"
            Range("ZA3").Value = FirstCat.Offset(0, 4).Value
            Set varTestRng = Range("ZA3")
            Set varTestStrt = Range("ZA3")
            Range("ZB3").Value = FirstCat.Offset(0, 2).Value
            If Not Range("ZB3").Value = SuppQty Then
                Range("ZB4").Value = SuppQty
            End If
            Set frstUOM = Range("ZB3")
            SubAtsFlg = 1
            Exit Sub
        End If
        
        
End Sub
Function FUN_suppChk(chkNm As String, Rw As Integer) As Integer

    'find contracted mfg section
    '-----------------------------
    On Error GoTo ERR_NoSupp
    FUN_suppChk = Rows("3:3").Find(what:=chkNm).Column
    If Range("X" & Rw).Offset(0, FUN_suppChk - Range("X" & Rw).Column) = "-" Then
        FUN_suppChk = 0
    Else
        'set pricefile variable and target minval if exists
        '------------------------------------
        Set FirstCon = FirstCat.Offset(0, FUN_suppChk - FirstCat.Column)
        conpos = (FirstCon.Column - Range("BG1").Column + 30) / 30
        SuppQty = FirstCon.Offset(0, 2).Value
        SuppDesc = FirstCon.Offset(0, 3).Value
        If HansEAval = 0 Then
            If Not PFstdFLG = 1 Then
                If SuppQty = 1 And (SuppDesc = "BN" Or SuppDesc = "CA" Or SuppDesc = "PL" Or SuppDesc = "BG" Or SuppDesc = "BX" Or SuppDesc = "CT" Or SuppDesc = "DZ" Or SuppDesc = "PK") Then
                    SuppMinVal = 0
                Else
                    SuppMinVal = FirstCon.Offset(0, 5).Value
                End If
            Else
                SuppMinVal = FirstCon.Offset(0, 5).Value
            End If
        End If
    End If
   
Exit Function
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ERR_NoSupp:
FUN_suppChk = 0
Exit Function


End Function

Sub ContextCheck()

'find UOMs in first 2 and aft 2
'create an array of prev 2 and next 2 plus all the catnums prev and aft until there's a different UOM than the UOMs in the first/next 2, create flg so if don't want to go on indefinitely then will only check surrounding 2
'make confidence level depending on how many of each UOM qty are found
'? include rvals for each UOM and pkgstr for each?

'maybe search all items from that mfg within rng of context R
'maybe just the adjoinging mfg items
'return an array of all UOMs with frequency of occurence
'then end there and analyze from main sub

dataCnt = 0
For Each c In Range(Range("N3"), Range("N2").End(xlDown))
    If c.Offset(0, -2).Value = ContextMfg Then
        If c.Offset(0, 4).Value > 0 Then
            Rval = c.Offset(0, 4).Value
            If (Rval - ContextRval) / Rval < plvlVarRange And (Rval - ContextRval) / Rval > -1 * plvlVarRange Then
                'Debug.Print c.Address
                'Debug.Print Rval
                pval = c.Offset(0, 2).Value
                If pval > 0 Then
                    dataCnt = dataCnt + 1
                    If Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown)), pval) > 0 Then
                        Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown)).Find(what:=pval, lookat:=xlWhole, LookIn:=xlFormulas).Offset(0, 1).Select
                        ActiveCell.Value = ActiveCell.Value + 1
                    Else
                        ReturnStrt.Offset(-2, 0).End(xlDown).Offset(1, 0).Value = pval
                        ReturnStrt.Offset(-2, 0).End(xlDown).Offset(0, 1).Value = 1
                    End If
                    If dataCnt = 100 Then GoTo contextend '<--exit if you have enough context data
                End If
            End If
        End If
    End If
Next

contextend:
If Not ReturnStrt.Value = "" Then

'    returnSet = Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown))
'    'Sort descending
'    '------------------------------------
'    For i = returnSet.Count - 1 To 1
'        sortminval = WorksheetFunction.Max(Range(returnSet(i), returnSet(i).End(xlDown)))
'        sortassoc =
'        Range(returnSet(i), returnSet(i).End(xlUp)).Find(what:=WorksheetFunction.Min(Range(returnSet(i), returnSet(i).End(xlUp))), lookat:=xlWhole).Value = returnSet(i).Value
'        returnSet(i).Value = sortminval
'    Next

    Set ReturnSet = Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown))
    Exit Sub
    
End If

'if no context then find context of upper and lower items:
'------------------------------------------

'Range("AAA:AAF").ClearContents
'Range("AAA1:AAF2").Value = "x"
''On Error GoTo 0
'
'If Not Range("N:N").Find(what:=FirstCat.Offset(-1, 0).Value, lookat:=xlWhole).Address = Range("N2").Address Then
'    Set prevfirstcat1 = Range("N:N").Find(what:=FirstCat.Offset(-1, 0).Value, lookat:=xlWhole)
'    prevmtchcats1 = Application.CountIf(Range("N:N"), prevfirstcat1.Value)
'    Set prevallcats1 = Range(prevfirstcat1, prevfirstcat1.Offset(prevmtchcats1 - 1, 0))
'Else
'    Set prevfirstcat1 = Range("N3").Address
'    prevmtchcats1 = 1
'End If
'
'If Not Range("N:N").Find(what:=prevfirstcat1.Offset(-1, 0).Value, lookat:=xlWhole).Address = Range("N2").Address Then
'    Set prevfirstcat2 = Range("N:N").Find(what:=prevfirstcat1.Offset(-1, 0).Value, lookat:=xlWhole)
'    prevmtchcats2 = Application.CountIf(Range("N:N"), prevfirstcat2.Value)
'    Set prevallcats2 = Range(prevfirstcat2, prevfirstcat2.Offset(prevmtchcats2 - 1, 0))
'Else
'    Set prevfirstcat2 = prevfirstcat1
'    prevmtchcats2 = prevmtchcats1
'End If
'
'If Not Trim(FirstCat.Offset(mtchcats, 0).Value) = "" And Not FirstCat.Offset(mtchcats, 0).Value = 0 Then
'    Set nextfirstcat1 = Range("N:N").Find(what:=FirstCat.Offset(mtchcats, 0).Value, lookat:=xlWhole)
'    nextmtchcats1 = Application.CountIf(Range("N:N"), nextfirstcat1.Value)
'    Set nextallcats1 = Range(nextfirstcat1, nextfirstcat1.Offset(nextmtchcats1 - 1, 0))
'Else
'    Set nextfirstcat1 = FirstCat
'    nextmtchcats1 = mtchcats
'End If
'
'If Not Trim(nextfirstcat1.Offset(nextmtchcats1, 0).Value) = "" And Not nextfirstcat1.Offset(nextmtchcats1, 0).Value = 0 Then
'    Set nextfirstcat2 = Range("N:N").Find(what:=nextfirstcat1.Offset(nextmtchcats1, 0).Value, lookat:=xlWhole)
'    nextmtchcats2 = Application.CountIf(Range("N:N"), nextfirstcat2.Value)
'    Set nextallcats2 = Range(nextfirstcat2, nextfirstcat2.Offset(nextmtchcats2 - 1, 0))
'Else
'    Set nextfirstcat2 = nextfirstcat1
'    nextmtchcats2 = nextmtchcats1
'End If
'
'For Each c In Range(FirstCat.Offset(-1, 0), prevfirstcat2)
'    Range("aaa1").End(xlDown).Offset(1, 0).Value = c.Offset(0, 1).Value     'UOM
'    Range("aaa1").End(xlDown).Offset(0, 1).Value = c.Offset(0, 3).Value     'pkgstr
'    Range("aaa1").End(xlDown).Offset(0, 2).Value = c.Offset(0, 4).Value     'rval
'    Range("aaa1").End(xlDown).Offset(0, 3).Value = c.Offset(0, 8).Value     'EAval
'Next
'For Each c In Range(FirstCat.Offset(mtchcats, 0), nextfirstcat2.Offset(nextmtchcats2 - 1, 0))
'    Range("aaa1").End(xlDown).Offset(1, 0).Value = c.Offset(0, 1).Value     'UOM
'    Range("aaa1").End(xlDown).Offset(0, 1).Value = c.Offset(0, 3).Value     'pkgstr
'    Range("aaa1").End(xlDown).Offset(0, 2).Value = c.Offset(0, 4).Value     'rval
'    Range("aaa1").End(xlDown).Offset(0, 3).Value = c.Offset(0, 8).Value     'EAval
'Next
'
'For Each c In Range(Range("AAA3"), Range("AAA1").End(xlDown))
'    If c.Offset(0, 2).Value < contextRval * 1.5 And c.Offset(0, 2).Value > contextRval * 0.5 Then ' / UOM.Value < MinVal + MinVal * plvlVarRange And Rval / UOM.Value > MinVal - MinVal * plvlVarRange) And Abs(Rval / UOM.Value - MinVal) < distVal Then
'        Range("AAE1").End(xlDown).Offset(1, 0).Value = c.Value
'        'c.offset(0,4),value = c.value
'    End If
'Next
''If Application.CountA(Range(Range("AAA1"), Range("AAA1").End(xlDown)).Offset(0, 4)) > 0 Then
'    ContextUOM = 0
'    For Each c In Range(Range("AAE3"), Range("AAE1").End(xlDown))
'        If Application.CountIf(Range(Range("AAE3"), Range("AAE1").End(xlDown)), c.Value) > Application.CountIf(Range(Range("AAE3"), Range("AAE1").End(xlDown)), ContextUOM) Then ContextUOM = c.Value
'    Next
'    ''Range("AAF1").Value = "=Mode(" & Range(Range("AAE3"), Range("AAE1").End(xlDown)).Address & ")"
''Else
''    ContextUOM = 0
''End If
'
'
End Sub
Sub GlobalContext()


dataCnt = 0

sheet(contextDesc).Select



For Each c In Range(Range("N3"), Range("N2").End(xlDown))
    If c.Offset(0, -2).Value = ContextMfg Then
        If c.Offset(0, 4).Value > 0 Then
            Rval = c.Offset(0, 4).Value
            If (Rval - ContextRval) / Rval < plvlVarRange And (Rval - ContextRval) / Rval > -1 * plvlVarRange Then
                'Debug.Print c.Address
                'Debug.Print Rval
                pval = c.Offset(0, 2).Value
                If pval > 0 Then
                    dataCnt = dataCnt + 1
                    If Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown)), pval) > 0 Then
                        Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown)).Find(what:=pval, lookat:=xlWhole, LookIn:=xlFormulas).Offset(0, 1).Select
                        ActiveCell.Value = ActiveCell.Value + 1
                    Else
                        ReturnStrt.Offset(-2, 0).End(xlDown).Offset(1, 0).Value = pval
                        ReturnStrt.Offset(-2, 0).End(xlDown).Offset(0, 1).Value = 1
                    End If
                    If dataCnt = 100 Then GoTo contextend '<--exit if you have enough context data
                End If
            End If
        End If
    End If
Next

contextend:
If Not ReturnStrt.Value = "" Then

'    returnSet = Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown))
'    'Sort descending
'    '------------------------------------
'    For i = returnSet.Count - 1 To 1
'        sortminval = WorksheetFunction.Max(Range(returnSet(i), returnSet(i).End(xlDown)))
'        sortassoc =
'        Range(returnSet(i), returnSet(i).End(xlUp)).Find(what:=WorksheetFunction.Min(Range(returnSet(i), returnSet(i).End(xlUp))), lookat:=xlWhole).Value = returnSet(i).Value
'        returnSet(i).Value = sortminval
'    Next

    Set ReturnSet = Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown))
    Exit Sub
    
End If

Exit Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlNOsht:
For Each sht In UOMwb.Sheets
    mtchcats = Application.CountIf(Range(sht.Range("B1"), sht.Range("B1").End(xlDown)), Catnmbr)
    If mtchcats > 0 Then
        
        
        Do
            On Error GoTo errhndlNOstd
            FoundVal = UOMwb.Sheets(shtnm).Range("A1:" & mtchColPrev & 10).Find(what:=ActiveCell.Offset(0, 1).Value, lookat:=xlPart).Address
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(1, 0).Value = FoundVal
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 1).Value = ActiveCell.Value
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 2).Value = ActiveCell.Offset(0, 1).Value
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 3).Value = CoreXrefWB.Sheets(shtnm).Range(FoundVal).Value
            CoreXrefWB.Sheets(shtnm).Range(FoundVal).ClearContents
            On Error GoTo errhndlNXTstd
NXTstdmfg:  FoundVal = CoreXrefWB.Sheets(shtnm).Range("A1:" & mtchColPrev & 10).Find(what:=ActiveCell.Offset(0, 1).Value, lookat:=xlPart).Address
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(1, 0).Value = FoundVal
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 1).Value = ActiveCell.Value
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 2).Value = ActiveCell.Offset(0, 1).Value
            CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 3).Value = CoreXrefWB.Sheets(shtnm).Range(FoundVal).Value
            CoreXrefWB.Sheets(shtnm).Range(FoundVal).ClearContents
            GoTo NXTstdmfg
'            For Each stdmfg In CoreXrefWB.Sheets(shtnm).Range("A1:" & mtchColPrev & 10)
'                If InStr(LCase(stdmfg), LCase(ActiveCell.Offset(0, 1).Value)) > 0 Then
'                    CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(1, 0).Value = stdmfg.Address
'                    CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 1).Value = ActiveCell.Value
'                    CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 2).Value = ActiveCell.Offset(0, 1).Value
'                    'CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 4).Value = ActiveCell.Value
'                End If
'            Next
NOstd:      For Each c In Range(ActiveCell, ActiveCell.Offset(Application.CountIf(Range("B:B"), ActiveCell.Offset(0, 1).Value) - 1, 0))
                On Error GoTo errhndlNOMFG
                FoundVal = CoreXrefWB.Sheets(shtnm).Range("A1:" & mtchColPrev & 10).Find(what:=c.Value, lookat:=xlPart).Address
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(1, 0).Value = FoundVal
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 1).Value = c.Value
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 2).Value = c.Offset(0, 1).Value
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 3).Value = CoreXrefWB.Sheets(shtnm).Range(FoundVal).Value
                CoreXrefWB.Sheets(shtnm).Range(FoundVal).ClearContents
                'c.Offset(Application.CountIf(Range("B:B"), c.Offset(0, 1).Value) - Application.CountIf(Range(c.Offset(0, 1), c.Offset(100, 1)), c.Offset(0, 1).Value), 0).Select
                On Error GoTo errhndlNXTstd
FndOtr:         FoundVal = CoreXrefWB.Sheets(shtnm).Range("A1:" & mtchColPrev & 10).Find(what:=c.Value, lookat:=xlPart).Address
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(1, 0).Value = FoundVal
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 1).Value = c.Value
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 2).Value = c.Offset(0, 1).Value
                CoreXrefWB.Sheets(shtnm).Range(mtchsupprng.Address).End(xlDown).Offset(0, 3).Value = CoreXrefWB.Sheets(shtnm).Range(FoundVal).Value
                CoreXrefWB.Sheets(shtnm).Range(FoundVal).ClearContents
                'c.Offset(Application.CountIf(Range("B:B"), c.Offset(0, 1).Value) - Application.CountIf(Range(c.Offset(0, 1), c.Offset(100, 1)), c.Offset(0, 1).Value), 0).Select
                GoTo FndOtr
NOmfg:      Next
NXTstd:     ActiveCell.Offset(Application.CountIf(Range("B:B"), ActiveCell.Offset(0, 1).Value), 0).Select
NXTmfg: Loop Until IsEmpty(ActiveCell.Offset(1, 0))
        
        
        'for each c in
        If (Rval - ContextRval) / Rval < plvlVarRange And (Rval - ContextRval) / Rval > -1 * plvlVarRange Then
            'Debug.Print c.Address
            'Debug.Print Rval
            pval = c.Offset(0, 2).Value
            If pval > 0 Then
                dataCnt = dataCnt + 1
                If Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown)), pval) > 0 Then
                    Range(ReturnStrt, ReturnStrt.Offset(-2, 0).End(xlDown)).Find(what:=pval, lookat:=xlWhole, LookIn:=xlFormulas).Offset(0, 1).Select
                    ActiveCell.Value = ActiveCell.Value + 1
                Else
                    ReturnStrt.Offset(-2, 0).End(xlDown).Offset(1, 0).Value = pval
                    ReturnStrt.Offset(-2, 0).End(xlDown).Offset(0, 1).Value = 1
                End If
                If dataCnt = 100 Then GoTo contextend '<--exit if you have enough context data
            End If
        End If
    End If
Next
    


End Sub
Sub Hans_METH()

HansFLG = 1
Call StrtEinstein

'Check setup is correct and Set values for next loop
'====================================================================================

'set catvals
'---------------------------------
Catnmbr = ActiveCell.Offset(0, Range("N1").Column - ActiveCell.Column).Value
Set FirstCat = Range("N:N").Find(what:=Catnmbr, lookat:=xlWhole)
mtchcats = Application.CountIf(Range("N:N"), FirstCat.Value)
ReDim Xarray(1 To mtchcats)
Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))

'check sort
'---------------------------------
endFLG = 0
Call SortChk
If endFLG = 1 Then
    endFLG = 0
    Exit Sub
End If

'Reset to original values
'---------------------------------
For Each c In AllCats
    c.Offset(0, 2).Value = c.Offset(0, 1).Value
Next
AllCats.Offset(0, 9).ClearContents
AllCats.Offset(0, 9).Interior.ColorIndex = 0
AllCats.Offset(0, 8).Interior.ColorIndex = 0

'check for form input
'====================================================================================

'EA values
'-------------------------
If Not ISNULL(RubiksForm.TargEAbx.Value) And RubiksForm.TargEAbx.Value <> "" And RubiksForm.TargEAbx.Value <> " " And RubiksForm.TargEAbx.Value <> "  " Then
    HansEAval = RubiksForm.TargEAbx.Value
Else
    HansEAval = 0
End If

'UOMs
'-------------------------
If Not ISNULL(RubiksForm.AddUOMbx.Value) And RubiksForm.AddUOMbx.Text <> "" And RubiksForm.AddUOMbx.Text <> " " And RubiksForm.AddUOMbx.Text <> "  " Then
    HansUOMflg = 1
    parskey = RubiksForm.AddUOMbx.Text
    Range(Range("YW3"), Range("YW1").End(xlDown)).ClearContents
    Range("YW1:YW2").Value = "x"
    If InStr(parskey, ";") > 0 Then
    
        'count number of ;
        '----------------------
        chrcount = 0
        For c = 1 To Len(parskey)
            If Mid(parskey, c, 1) = ";" Then
                chrcount = chrcount + 1
            End If
        Next
        
        'Parse UOMs
        '----------------------
        parscount = 0
        Do
            parscount = parscount + 1
            On Error GoTo errhndlParskey
            Range("YW2").Offset(parscount, 0).Value = Trim(Left(parskey, InStr(parskey, ";") - 1))
            parskey = Mid(parskey, InStr(parskey, ";") + 1, 100)
            On Error GoTo 0
        Loop Until parscount = chrcount + 1
    Else
        Range("YW3").Value = Trim(parskey)
    End If
    Set HansSet = Range(Range("YW3"), Range("YW1").End(xlDown))

Else
    Range(Range("YW3"), Range("YW1").End(xlDown)).ClearContents
    HansUOMflg = 0
End If

'run Einstein
'====================================================================================
FirstCat.Select
Call CommonSetup  '>>>>>>>>>>
ttlCurrVar = 0
Call VarCheck
If Not ttlCurrVar = ttlPrevVar Then Call MainLogic

'Clean up
'====================================================================================
'HansFLG = 0
HansEAval = 0
HansUOMflg = 0
AllCats.EntireRow.Calculate
Application.ScreenUpdating = True
AllCats.EntireRow.Select
Call FUN_CalcBackOn

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhndlParskey:
Range("YW2").Offset(parscount, 0).Value = Trim(parskey)
Resume Next


End Sub
Sub SortChk()

    'check to make sure sorted by catalog number
    '---------------------------------
    For Each c In AllCats
        If c.Value = FirstCat.Value Then
            contigchk = contigchk + 1
        End If
    Next
    If Not contigchk = mtchcats Then
        MsgBox "Please make sure your sheet is sorted by catalog number and try again."
        endFLG = 1
        Exit Sub
    End If
    
    'check to make sure sorted by R ascending
    '---------------------------------
    For Each c In Range(FirstCat, FirstCat.Offset(mtchcats - 2, 0))
        If c.Offset(0, 6).Value <= c.Offset(1, 6).Value Then
            Rchk = Rchk + 1
        End If
    Next
    If Not Rchk = mtchcats - 1 Then Call FUN_Sort("Line Item Data", AllCats.EntireRow, FirstCat.Offset(0, 6), 1)


End Sub
Sub CheckXrefs()

xrefCNT = 0
For i = 0 To suppNMBR - 1
    If Not i = conpos And Range("AO" & FirstCat.Row).Offset(0, i * 18).Value = False Then
        xrefCNT = xrefCNT + 1
        ReDim Preserve xrefArray(1 To xrefCNT)
        xrefArray(xrefCNT) = Range("AM" & FirstCat.Row).Offset(0, i * 18).Address
        xrefQty = Range("AP" & FirstCat.Row).Offset(0, i * 18).Value
        If Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), xrefQty) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = xrefQty
    End If
Next


End Sub
Sub CheckBenches()

BenchCNT = 0
For i = 0 To suppNMBR - 1
    If Not Trim(Range("HT" & FirstCat.Row).Offset(0, i * 16).Value) = "" Then 'And Not Range("AN" & FirstCat.Row).Offset(0, I * 18).Value = FirstCat.Value Then
        BenchCNT = BenchCNT + 1
        ReDim Preserve BenchArray(1 To BenchCNT)
        BenchArray(BenchCNT) = Range("HT" & FirstCat.Row).Offset(0, i * 16).Address
        BenchQty = Sheets("Best Market Price").Range("A:A").Find(what:=Range("AN" & FirstCat.Row).Offset(0, i * 18).Value, lookat:=xlWhole).Offset(0, 10).Value
        If Not Application.CountIf(Range(frstUOM, Range("ZA1").End(xlToRight).End(xlDown)), BenchQty) > 0 Then Range("ZA1").End(xlToRight).End(xlDown).Offset(1, 0).Value = BenchQty
    End If
Next


End Sub
Function FUN_VarVal(highEAval As Variant, lowEAval As Variant)

If highEAval = 0 Or highEAval = "-" Or highEAval = "" Or lowEAval = 0 Then
    FUN_VarVal = "-"
Else
    FUN_VarVal = (highEAval - lowEAval) / highEAval
'FUN_VarVal = (lowEAval - highEAval) / lowEAval
End If



End Function
