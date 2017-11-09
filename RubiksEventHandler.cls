VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RubiksEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents XEvents As MSForms.ToggleButton
Attribute XEvents.VB_VarHelpID = -1
Public WithEvents NavEvents As MSForms.CommandButton
Attribute NavEvents.VB_VarHelpID = -1
Public WithEvents CollEvents As MSForms.ToggleButton
Attribute CollEvents.VB_VarHelpID = -1
Public WithEvents UOMEvents As MSForms.TextBox
Attribute UOMEvents.VB_VarHelpID = -1
Dim CtrlRw As Integer
Dim currUOMevents
Dim UOM As String
Dim EA_Price As String
Dim Var As String
Private Sub XEvents_click()

    Dim currEA As Double
    Dim lowestEA As Double

    'change xval
    '--------------------------------------------
    If XEvents = True Then
        XEvents.Picture = LoadPicture(IconsPATH & "\" & XoutIcon, XEvents.Width + 2, XEvents.Height + 2)
        'XEvents.BackColor = &H0&
        XEvents = True
    Else
        XEvents.Picture = Nothing
        'XEvents.BackColor = &H8000000F
        XEvents = False
    End If
    
    'if x'd val is lowest EA then recalculate
    '--------------------------------------------
    recalcFlg = True
    CtrlRw = Replace(XEvents.Name, Section(1) & DataCtrl(7) & 1, "")
    lowestEA = RubiksForm.Controls(Section(1) & DataCtrl(6) & 1 & CtrlRw).Caption
    For j = 1 To mtchcats
        currEA = RubiksForm.Controls(Section(1) & DataCtrl(6) & 1 & j).Caption
        If currEA <= lowestEA Then
            If Not RubiksForm.Controls(Section(1) & DataCtrl(7) & 1 & j).Value = True And Not j = CtrlRw Then recalcFlg = False
        End If
    Next

    'recalculate variances
    '--------------------------------------------
    If Not recalcFlg = False Then
        Set currUOMevents = RubiksForm.Controls(Section(1) & DataCtrl(2) & 1 & CtrlRw)
        Call handleUOMchg(1)
    End If
    
    'Check if any high Vars or Nulls in other sections
    '------------------------------
    For sectn = 1 To 4
        If sectn = 1 Then
            Call VarFormatting(RubiksForm.Controls(Section(1) & DataCtrl(8) & 1 & CtrlRw), plvlVarRange, CtrlRw)
        ElseIf sectn = 2 Then
            If Not Sheets("line item data").Range("AR" & FirstCat.Row).Value = "-" Then Call VarFormatting(RubiksForm.Controls(Section(2) & DataCtrl(8) & 1 & CtrlRw), BnchVarRange, CtrlRw)
        ElseIf sectn = 3 Then
            For mftr = 1 To suppNMBR
                If Not Sheets("line item data").Range("BG" & FirstCat.Row).Offset(0, (mftr - 1) * 30).Value = "-" Then
                    Call VarFormatting(RubiksForm.Controls(Section(3) & DataCtrl(8) & mftr & CtrlRw), suppVarRange, CtrlRw)
                End If
            Next
        ElseIf sectn = 4 Then
            For mftr = 1 To suppNMBR
                If Not Sheets("line item data").Range("BV" & FirstCat.Row).Offset(0, (mftr - 1) * 30).Value = "" Then
                    Call VarFormatting(RubiksForm.Controls(Section(4) & DataCtrl(8) & mftr & CtrlRw), BnchVarRange, CtrlRw)
                End If
            Next
        End If
    Next
    'End If



End Sub
Private Sub NavEvents_click()


Application.ScreenUpdating = True
    
    mfgpos = FUN_MfgPos(NavEvents.Parent.Name)
    If InStr(NavEvents.Name, Section(1)) Then
        Sheets("Line item data").Select
        AllCats.EntireRow.Select
    ElseIf InStr(NavEvents.Name, Section(2)) Then            '<<<Check this
        Call FUN_TestForSheet("Best market price")
        On Error Resume Next
        Range("A:A").Find(what:=Catnmbr, lookat:=xlWhole).EntireRow.Select
    ElseIf InStr(NavEvents.Name, Section(3)) > 0 And Not InStr(NavEvents.Name, Section(4)) > 0 Then
        suppAlias = FUN_SuppName(mfgpos)
        If NavEvents.BackColor = &H80FF80 Then
            'contracted supplier
            '---------------------
            Call FUN_TestForSheet(suppAlias & " Pricing")
            On Error Resume Next
            Range("A:A").Find(what:=Catnmbr, lookat:=xlWhole).EntireRow.Select
        Else
            'xref supplier
            '---------------------
            Call FUN_TestForSheet(suppAlias & " Cross Reference")
            ActiveSheet.AutoFilterMode = False
            ActiveSheet.Range("A1:O" & FUN_lastrow("A")).AutoFilter Field:=1, Criteria1:=Catnmbr
            ActiveWindow.ScrollRow = 1
        End If
    ElseIf InStr(NavEvents.Name, Section(4)) > 0 Then
        FUN_TestForSheet ("Best market price")
        On Error Resume Next
        BenchCatNmbr = Sheets("line item data").Range("BG" & FirstCat.Row).Offset(0, (mfgpos - 1) * 30).Value
        Range("A:A").Find(what:=BenchCatNmbr, lookat:=xlWhole).EntireRow.Select
    End If

End Sub
Private Sub CollEvents_click()


    Set ctrlframe = RubiksForm.Controls(Replace(CollEvents.Name, "Collapse", "Frame"))
'    offwdth = CtrlFrame.Width - CollEvents.Width
'    mfgpos = FUN_MfgPos(CollEvents.Name)

    'move selected controls
    '----------------------------------
'    CollEvents.Left = Evaluate(Evaluate(CollEvents.Left & op & CtrlFrame.Width) & op2 & CollEvents.Width)
'    CollEvents.Top = Evaluate(CollEvents.Top & op2 & CollEvents.Height)
    
    'move controls on the right of selected control
    '==================================================================================================================

        'Move frame and move collapse
        '---------------------------------
        If CollEvents = True Then
            ctrlframe.Visible = False
            oprtr = "+"
            oprtr2 = "-"
        Else
            ctrlframe.Visible = True
            oprtr = "-"
            oprtr2 = "+"
        End If
        
        'move section frame
        '---------------------------------
        ctrlframe.Left = Evaluate(ctrlframe.Left & oprtr2 & ctrlframe.Width & oprtr & CollEvents.Width)
        'ctrlframe.Top = Evaluate(ctrlframe.Top & oprtr & CollEvents.Height)
        CollEvents.Left = Evaluate(CollEvents.Left & oprtr2 & ctrlframe.Width & oprtr & CollEvents.Width)
        'CollEvents.Left = Evaluate(RubiksForm.Controls(Section(1) & "Frame" & 1).Width & oprtr2 & CtrlFrame.Width & oprtr & CollEvents.Width)
        CollEvents.Top = Evaluate(CollEvents.Top & oprtr & CollEvents.Height)
        
        
        'move frames to the right of section frame
        '---------------------------------
        For i = 1 To suppNMBR
            If Not Sheets("line item data").Range("BG" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "-" Then
                Set CurrFrame = RubiksForm.Controls(Section(3) & "Frame" & i)
                If CurrFrame.Left > ctrlframe.Left Or InStr(ctrlframe.Name, Section(2)) > 0 Then
                'If RubiksForm.Controls(Section(3) & "Collapse" & i).Left > CollEvents.Left Then
                    CurrFrame.Left = Evaluate(CurrFrame.Left & oprtr2 & ctrlframe.Width & oprtr & CollEvents.Width)
                    RubiksForm.Controls(Section(3) & "Collapse" & i).Left = Evaluate(CurrFrame.Left & "+" & CurrFrame.Width & "-" & CollEvents.Width)
                End If
                If Not Sheets("line item data").Range("BV" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "" Then
                    Set CurrFrame = RubiksForm.Controls(Section(4) & "Frame" & i)
                    If CurrFrame.Left > ctrlframe.Left Or InStr(ctrlframe.Name, Section(2)) > 0 Then
                    'If RubiksForm.Controls(Section(4) & "Collapse" & i).Left > CollEvents.Left Then
                        CurrFrame.Left = Evaluate(CurrFrame.Left & oprtr2 & ctrlframe.Width & oprtr & CollEvents.Width)
                        RubiksForm.Controls(Section(4) & "Collapse" & i).Left = Evaluate(CurrFrame.Left & "+" & CurrFrame.Width & "-" & CollEvents.Width)
                    End If
                End If
            End If
        Next


    If Evaluate(RubiksForm.Width & oprtr2 & ctrlframe.Width) < RubiksForm.CIAbttn.Left + RubiksForm.CIAbttn.Width Then
        RubiksForm.Width = RubiksForm.CIAbttn.Left + RubiksForm.CIAbttn.Width + 4
    ElseIf IsEmpty(CurrFrame) Then
        RubiksForm.Width = Evaluate(RubiksForm.Controls(Replace(ctrlframe.Name, "Frame", "Collapse")).Left + CollEvents.Width) + 4
    Else
        'RubiksForm.Width = Evaluate(RubiksForm.Width & oprtr2 & CtrlFrame.Width) + 3
        RubiksForm.Width = Evaluate(RubiksForm.Controls(Replace(CurrFrame.Name, "Frame", "Collapse")).Left + CollEvents.Width) + 4
    End If


End Sub
Sub UOMEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'For Each c In RubiksForm.HomeFrame.Controls
'    If InStr(LCase(c.Name), "mbr") Then Debug.Print c.Name
'Next

If KeyCode = 39 Then 'And Shift = 1 Then
    If Not InStr(LCase(UOMEvents.Name), LCase(Section(1) & DataCtrl(2))) > 0 Then Exit Sub
    On Error GoTo errhndlBOF
    Rw = Right(UOMEvents.Name, 1)
    If Rw = 9 Then
        If Right(UOMEvents.Name, 2) = 19 And Not IsNumeric(Right(UOMEvents.Name, 3)) Then
            cropLen = 1
        Else
            cropLen = 2
        End If
        Set currUOMevents = RubiksForm.Controls(Left(UOMEvents.Name, Len(UOMEvents.Name) - cropLen) & Right(UOMEvents.Name, cropLen) + 1)
    Else
        Set currUOMevents = RubiksForm.Controls(Left(UOMEvents.Name, Len(UOMEvents.Name) - 1) & Rw + 1)
    End If
    On Error GoTo 0
    currUOMevents.Text = UOMEvents.Text
    currUOMevents.SetFocus
    Call handleUOMchg(0)
    
'    'run again for current so you don't have to go back and do it again
'    '--------------------------------
'    Set currUOMevents = UOMEvents
'    Call handleUOMchg
    
ElseIf Not KeyCode = 13 And Not KeyCode = 9 Then
    Exit Sub  '(trigger event on Enter or Tab)
Else
    Set currUOMevents = UOMEvents
    Call handleUOMchg(0)
End If

Exit Sub
':::::::::::::::::::::::::::::
errhndlBOF:
Exit Sub


End Sub
Sub handleUOMchg(Optional Xcall)

Dim MinEA As Double
Dim currEA As Double

'<<defined in rubiks form>>
'---------------------------------
'DataCtrl(1) = Me.HomeOrgnl.Name
'DataCtrl(2) = Me.HomeUOM.Name
'DataCtrl(3) = Me.HomePkg.Name
'DataCtrl(4) = Me.HomeUCost.Name
'DataCtrl(5) = Me.HomeUsage.Name
'DataCtrl(6) = Me.HomeEPrice.Name
'DataCtrl(7) = Me.HomeX.Name
'DataCtrl(8) = Me.HomeVar.Name

    UOM = DataCtrl(2)
    EA_Price = DataCtrl(6)
    Var = DataCtrl(8)

'Find Section
'---------------------------------
If InStr(currUOMevents.Name, Section(1)) > 0 And Not InStr(currUOMevents.Name, Section(2)) > 0 Or Xcall = 1 Then

    mbrEA_CtrlStr = Section(1) & EA_Price & 1
    X_CtrlStr = Section(1) & DataCtrl(7) & 1
    MbrVar_CtrlStr = Section(1) & Var & 1
    
    If Xcall = 1 Then
        itmRw = Replace(currUOMevents.Name, X_CtrlStr, "")
    Else
        'Change UOM
        '------------------------------
        itmRw = Replace(currUOMevents.Name, Section(1) & DataCtrl(2) & 1, "")
        RubiksForm.Controls(Section(1) & DataCtrl(2) & "1" & itmRw).Text = currUOMevents.Text
        
        'Change EAval
        '------------------------------
        EAval = FUN_EAval(Section(1), 1, itmRw)
        RubiksForm.Controls(mbrEA_CtrlStr & itmRw).Caption = EAval
        
    End If
    
    'calculate plvling variance
    '------------------------------
    MinEA = 1000000
    For i = 1 To mtchcats
        currEA = RubiksForm.Controls(mbrEA_CtrlStr & i).Caption
        If currEA < MinEA Then
            If Not RubiksForm.Controls(X_CtrlStr & i).Value = True Then MinEA = currEA
            'If RubiksForm.Controls(X_CtrlStr & i).img = "" Then MinEA = currEA
        End If
    Next
    For i = 1 To mtchcats
        RubiksForm.Controls(MbrVar_CtrlStr & i).Caption = FUN_VarVal(RubiksForm.Controls(mbrEA_CtrlStr & i).Caption, MinEA)
        Call VarFormatting(RubiksForm.Controls(MbrVar_CtrlStr & i), plvlVarRange, i)
    Next

    If Not Xcall = 1 Then
        
        'calculate MbrBench and each Supp variance
        '------------------------------
        For i = 1 To suppNMBR
            If Not Range("BG" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "-" Then Call RecalculateVars_All(3, i)  '>>>>>>>>>>>>
            If Not Range("BV" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "" Then Call RecalculateVars_All(4, i)  '>>>>>>>>>>>>
        Next
        If Not Range("AR" & FirstCat.Row).Value = "-" Then Call RecalculateVars_All(2, 1)      '>>>>>>>>>>>>
        
    End If
    
        
ElseIf InStr(currUOMevents.Name, Section(2)) Then
    
    Call UpdateSectionVals(2, 1)  '>>>>>>>>>>
    
    'check if any suppBenchs are tied to mbrBench
    '----------------------------------------------
    For i = 1 To suppNMBR
        If Not Range("BV" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "" And FUN_BenchCatCompare(i) = True Then
            For j = 1 To mtchcats
                Call UpdateSectionVals(4, i) '>>>>>>>>>>
            Next
        End If
    Next
    
ElseIf InStr(currUOMevents.Name, Section(3)) > 0 And Not InStr(currUOMevents.Name, Section(4)) > 0 Then
    
    mfgpos = FUN_MfgPos(currUOMevents.Parent.Name)
    Call UpdateSectionVals(3, mfgpos) '>>>>>>>>>>
    
ElseIf InStr(currUOMevents.Name, Section(4)) > 0 Then
    
    mfgpos = FUN_MfgPos(currUOMevents.Parent.Name)
    Call UpdateSectionVals(4, mfgpos) '>>>>>>>>>>
    
    'check if MbrBench is tied to SuppBench
    '----------------------------------------------
    If Not FirstCat.Offset(0, 20).Value = "-" And FUN_BenchCatCompare(mfgpos) = True Then Call UpdateSectionVals(2, 1)
    
End If



End Sub
Sub UpdateSectionVals(sectn, mfgpos)

    'change MbrBench UOM
    '----------------------------------------------
    RubiksForm.Controls(Section(sectn) & DataCtrl(2) & mfgpos & 1).Text = currUOMevents.Text

    'change EAval
    '----------------------------------------------
    RubiksForm.Controls(Section(sectn) & DataCtrl(6) & mfgpos & 1).Caption = FUN_EAval(Section(sectn), mfgpos, 1)
    
    'Recalculate Variances
    '----------------------------------------------
    If sectn = 2 Then
        
        'recalculate supp vs mbr
        '--------------------------------
        Call RecalculateVars_All(2, 1)
        
    ElseIf sectn = 3 Then
        
        'recalculate supp vs mbr
        '--------------------------------
        Call RecalculateVars_All(3, mfgpos)
        
        'recalculate suppbench vs supp
        '--------------------------------
        'Call RecalculateVars_One(mfgpos)
        
    ElseIf sectn = 4 Then
        
        'recalculate suppbench vs supp
        '--------------------------------
        Call RecalculateVars_All(4, mfgpos)
        
    End If


End Sub
Sub RecalculateVars_All(sectn, mfgpos)

    For i = 1 To mtchcats
        RubiksForm.Controls(Section(sectn) & DataCtrl(8) & mfgpos & i).Caption = FUN_VarVal(RubiksForm.Controls(Section(sectn) & DataCtrl(6) & mfgpos & 1), RubiksForm.Controls(Section(1) & DataCtrl(6) & 1 & i).Caption)
        Call VarFormatting(RubiksForm.Controls(Section(sectn) & DataCtrl(8) & mfgpos & i), BnchVarRange, i)
    Next

End Sub
Sub RecalculateVars_One(mfgpos)

    VarVal = FUN_VarVal(RubiksForm.Controls(Section(4) & DataCtrl(6) & mfgpos & 1), RubiksForm.Controls(Section(3) & DataCtrl(6) & mfgpos & 1).Caption)
    For i = 1 To mtchcats
        RubiksForm.Controls(Section(4) & DataCtrl(8) & mfgpos & i).Caption = VarVal
        Call VarFormatting(RubiksForm.Controls(Section(4) & DataCtrl(8) & mfgpos & i), BnchVarRange, i)
    Next
    
End Sub
Sub VarFormatting(VarCtrl, VarThreshold, itm)
    
    'Dim VarVal As Double
    
    'If VarCtrl.Caption = "-" Then
        'If RubiksForm.Controls(Section(1) & DataCtrl(7) & 1 & itm).Value = True Then VarCtrl.BackColor = &H8000000F
        'Exit Sub
    'End If
    
    If Not VarCtrl.Caption = "-" Then
        VarVal = Val(Format(VarCtrl.Caption, "0.0000"))
        If VarVal < -100 Or VarVal > 100 Then
            VarCtrl.Caption = Format(1, "0%")
        Else
            VarCtrl.Caption = Format(VarVal, "0%")
        End If
    End If
    
    'Backcolor
    '---------------
    If RubiksForm.Controls(Section(1) & DataCtrl(7) & 1 & itm).Value = True Then
        VarCtrl.BackColor = &H8000000F
    ElseIf VarCtrl.Caption = "-" Then
        VarCtrl.BackColor = &HFF&
    ElseIf VarVal >= VarThreshold Or VarVal <= -VarThreshold Then
        VarCtrl.BackColor = &HFF&
    Else
        VarCtrl.BackColor = &H8000000F
    End If


End Sub
Function FUN_EAval(sectn, mfgpos, itmRw)


    Set EApriceCtrl = RubiksForm.Controls(sectn & DataCtrl(6) & mfgpos & itmRw)
    UOMcostCtrl = RubiksForm.Controls(sectn & DataCtrl(4) & mfgpos & itmRw).Caption
            
    If Not currUOMevents.Text = 0 And Not currUOMevents.Text = "" And Not UOMcostCtrl = "" Then
        FUN_EAval = Round(UOMcostCtrl / currUOMevents.Text, 6)
        If Left(FUN_EAval, 1) = 0 And Len(FUN_EAval) > 4 And Not FUN_EAval = 0 Then
            FUN_EAval = Format(FUN_EAval, "$0.0000")
        Else
            FUN_EAval = Format(FUN_EAval, "$0.00")
        End If
        'FUN_EAval = Format(UOMcostCtrl / currUOMevents.Text, "$0.00")
'        If FUN_EAval = 0 Then
'            epricectrl.BackColor = &HFF&
'        Else
'            epricectrl.BackColor = &H8000000F
'        End If
    Else
        FUN_EAval = Format(0, "$0.00")
    End If
    
        
End Function
Function FUN_BenchCatCompare(mfgpos) As Boolean

If FirstCat.Value = Sheets("line item data").Range("BG" & FirstCat.Row).Offset(0, (mfgpos - 1) * 30).Value Then
    FUN_BenchCatCompare = True
Else
    FUN_BenchCatCompare = False
End If

End Function
Function FUN_MfgPos(ctrlStr)

'make sure you get the parent ctrl, parent ctrl will be the frame and will not have and itmnmbr at the end of the ctrlname
'*********************************

'prntCtrlStr = RubiksForm.Controls(ctrlStr).Parent.Name
If Not Right(ctrlStr, 1) = 0 Then
    FUN_MfgPos = Right(ctrlStr, 1)
Else
    FUN_MfgPos = 10
End If

End Function


