VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RubiksForm 
   Caption         =   "Rubiks"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   23730
   OleObjectBlob   =   "RubiksForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RubiksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************
'Naming convention for form controls: (Section & Ctrl & Supplier & Itm) ; ([Member/MbrBench/Supplier/SuppBench] & [Ctrl label] & [1-suppnmbr] & [1-itms]) ; Ex. (SupplierUOMData213)
'Good Luck
'***********************



Dim Xarray() As New RubiksEventHandler
Dim NavArray() As New RubiksEventHandler
Dim CollArray() As New RubiksEventHandler
Dim UOMtxtArray() As New RubiksEventHandler
'Dim DataArray(1 To 8) As String
Dim PrevHeight As Double
Dim RubiksBrowser As InternetExplorer
Dim HansCancelInitial As Integer
Dim CollapseIcon As String
Dim FormWdth As Double
Dim FrameWdth As Double
Dim PrevCat As Range

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub UserForm_Initialize()

'set common variables
'=============================================================================================================
SetupSwitch = FUN_SetupSwitch

'Setup variable arrays
'=============================================================================================================
ReDim Preserve NavArray(1 To 1)
Set NavArray(1).NavEvents = MemberNav1

Section(1) = "Member"
Section(2) = "MbrBench"
Section(3) = "Supplier"
Section(4) = "SuppBench"

'MSBTarray(3) = "Bench"
'MSBTarray(4) = "Ten"

'<<Each header array parallels each other>>
'DataCtrl(1) = Replace(Me.Orgnl.Name, "Hdr", "Data")
'DataCtrl(2) = Replace(Me.UOM.Name, "Hdr", "Data")
'DataCtrl(3) = Replace(Me.Pkg.Name, "Hdr", "Data")
'DataCtrl(4) = Replace(Me.UCost.Name, "Hdr", "Data")
'DataCtrl(5) = Replace(Me.Usage.Name, "Hdr", "Data")
'DataCtrl(6) = Replace(Me.EA_Price.Name, "Hdr", "Data")
'DataCtrl(7) = Replace(Me.Xout.Name, "Hdr", "Data")
'DataCtrl(8) = Replace(Me.Var.Name, "Hdr", "Data")

DataCtrl(1) = "OrgnlData"
DataCtrl(2) = "UOMData"
DataCtrl(3) = "PkgData"
DataCtrl(4) = "UcostData"
DataCtrl(5) = "UsageData"
DataCtrl(6) = "EAPriceData"
DataCtrl(7) = "XoutData"
DataCtrl(8) = "VarData"

'get icon paths
'------------------------
Set fileObj = objFSO.GetFile(AdminConfigPATH)
AdminconfigStr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)
IconStr = FUN_ConvGroups(AdminconfigStr, "Icons")
    IconsPATH = FUN_ConvTags(IconStr, "Local Folder")
    CollapseIcon = FUN_ConvTags(IconStr, "Collapse Icon")
    XoutIcon = FUN_ConvTags(IconStr, "Xout Icon")

'set variance range
'------------------------
'plvlvarTop = Sheets("Line Item Data").Rows("2:2").Find(what:="Price Level Annual Spend").Offset(0, -1).Address   '(HN2)
Set PlvlStrt = Sheets("Line Item Data").Rows("4:4").Find(what:="Price Level Each Cost")   '(AM4)
plvlCol = PlvlStrt.Column '(AM column number)
plvloffFromX = plvlCol - Sheets("Line Item Data").Range("X2").Column

Me.Width = MemberNav1.Width + 4 '+ RunCat.Width + 2
Me.Height = 43 'CatTitle.Height + MemberNav1.Height + HomeUOM.Height

MenuFrame.Visible = False
AddUOMframe.Visible = False
TargEAFrame.Visible = False

CatTitle.Text = vbNullString
ItemsLeft.Caption = vbNullString

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2

Call FUN_CalcOff

popoutFLG = 1

'AddToForm MIN_BOX


End Sub
Private Sub UserForm_Activate()

If HansCancelInitial = 1 Then
    Unload RubiksForm
Else
    AddToForm MIN_BOX
End If


End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'[TBD] make keyboard shortcut for everything so don't have to use mouse
'key <- and -> ^ v to move between UOMtxtboxes and xevents
'Shift+ -> <- to move between catnmbrs
'shift+A to accept
'Shift+[1-10] for supp navs
'Shift+B+[1-10] for bench nav
'Shift+M/T for Mbr/Ten navs
'Shift+L to turn on logic helper
'Shift+O to set back to orgnl vals
'Shift+N to add note

'add ctrls to show each section of HCO detail tab and impact summary, and each frequently used tab, like CMSgraph, and  like bkmrks
'but maybe add these as part of Zues navigation, or maybe just add to both
'have an option to link Hans and Zues through public variable so that if one moves then the other moves


End Sub
Private Sub RunCat_Click()

Dim FormHght As Double

Dim MemberValCol(1 To 8) As String
MemberValCol(1) = "AA"   '<<Orgnl
MemberValCol(2) = "AB"   '<<UOMqty
MemberValCol(3) = "AC"   '<<UOMpkg
MemberValCol(4) = "AD"   '<<UOM price
MemberValCol(5) = "AF"   '<<Annualized usage
MemberValCol(6) = "AH"   '<<EA price
MemberValCol(7) = "AI"   '<<Xout
MemberValCol(8) = "AO"   '<<plvlVar

Dim MbrBenchValCol(1 To 8) As String
MbrBenchValCol(1) = vbNullString
MbrBenchValCol(2) = "X"
MbrBenchValCol(3) = vbNullString
MbrBenchValCol(4) = "AR"
MbrBenchValCol(5) = vbNullString
MbrBenchValCol(6) = "AS"
MbrBenchValCol(7) = vbNullString
MbrBenchValCol(8) = "AV"

Dim SupplierValCol(1 To 8) As String
SupplierValCol(1) = vbNullString
SupplierValCol(2) = "BI"
SupplierValCol(3) = "BJ"
SupplierValCol(4) = "BK"
SupplierValCol(5) = vbNullString
SupplierValCol(6) = "BL"
SupplierValCol(7) = vbNullString
SupplierValCol(8) = "BN"

Dim SuppBenchValCol(1 To 8) As String
SuppBenchValCol(1) = vbNullString
SuppBenchValCol(2) = "BG"
SuppBenchValCol(3) = vbNullString
SuppBenchValCol(4) = "BV"
SuppBenchValCol(5) = vbNullString
SuppBenchValCol(6) = "BW"
SuppBenchValCol(7) = vbNullString
SuppBenchValCol(8) = "BZ"

'Setup
'==========================================================================================================
If Not ActiveSheet.Name = "Line Item Data" Then
    MsgBox "Please select catalog number on the Line Item Data tab"
    Exit Sub
End If

If LogicHelp = True Then
    
    Call Hans_METH  '>>>>>>>>>>
    
    If endFLG = 1 Then
        endFLG = 0
        Exit Sub
    End If

Else
    plvlVarRange = Val(Format(ZeusForm.plvlRngSet.Value, "0.00"))
    suppVarRange = Val(Format(ZeusForm.SuppRngSet.Value, "0.00"))
    BnchVarRange = Val(Format(ZeusForm.bnchRngSet.Value, "0.00"))
    
    'set catvals
    '---------------------------------
    Catnmbr = ActiveCell.Offset(0, Range("X1").Column - ActiveCell.Column).Value
    Set PrevCat = FirstCat
    Set FirstCat = Range("X:X").Find(what:=Catnmbr, lookat:=xlWhole)
    mtchcats = Application.CountIf(Range("X:X"), FirstCat.Value)                  'Count number of matching catalog numbers
    Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))
    ConCol = FUN_suppChk(FirstCat.Offset(0, -2).Value, FirstCat.Row)

    'check sort
    '---------------------------------
    endFLG = 0
    Call SortChk
    If endFLG = 1 Then
        endFLG = 0
        Exit Sub
    End If

End If

Call Calculate_Priceleveling_Single(AllCats.Offset(0, 15)) '>>>>>>>>>>

CatTitle.Text = FirstCat.Value
Me.Controls(Section(1) & "Frame" & 1).Visible = True
WebRefNote = vbNullString

If Sheets("best Market Price").AutoFilterMode = True Then
    Sheets("best Market Price").AutoFilterMode = False
    Sheets("best Market Price").Range("A1:P1").AutoFilter
End If

'clear previous data
'=============================================================================================================
On Error Resume Next
For Each clrCtrl In Me.Controls
    If TypeName(clrCtrl) = "Frame" Then
        If (InStr(clrCtrl.Name, Section(2)) Or InStr(clrCtrl.Name, Section(3)) Or InStr(clrCtrl.Name, Section(4))) Then Me.Controls.Remove clrCtrl.Name
    ElseIf InStr(clrCtrl.Name, Section(1)) Then
        If InStr(clrCtrl.Name, "Data") > 0 Then Me.Controls.Remove clrCtrl.Name
    ElseIf InStr(clrCtrl.Name, "Collapse") Then
        Me.Controls.Remove clrCtrl.Name
    End If
Next
On Error GoTo 0

'Populate data for Supp,Bench,Ten pricing
'(I nested all the loops and used skips to ensure consistency in frame/control parameters, if I have to change a parameter I only have to change it once for all sections/headers/labels/ect.)
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormWdth = Me.Controls(Section(1) & "Frame" & 1).Width
For Ctrl = 1 To UBound(DataCtrl)
    If Me.Controls(Section(1) & Replace(DataCtrl(Ctrl), "Data", "Hdr") & 1).Height > MaxhdrHght Then MaxhdrHght = Me.Controls(Section(1) & Replace(DataCtrl(Ctrl), "Data", "Hdr") & 1).Height
Next
FrameHght = Me.Controls(Section(1) & "Nav" & 1).Height + MaxhdrHght + mtchcats * MaxhdrHght + 7

'populate additional frames and data
'===========================================================================================================================================
For sectn = 1 To 4
    
    If sectn = 2 And Range("AR" & FirstCat.Row).Value = "-" Then GoTo SectionSkip
    If sectn = 1 Or sectn = 2 Then
        frames = 1
    Else
        frames = suppNMBR
    End If
    
    'Section Setup
    '===========================================================================================================================================
    For mfgpos = 1 To frames
    
        If sectn = 3 And Range("BG" & FirstCat.Row).Offset(0, (mfgpos - 1) * 30).Value = "-" Then GoTo FrameSkip
        If sectn = 4 And Range("BV" & FirstCat.Row).Offset(0, (mfgpos - 1) * 30).Value = "" Then GoTo FrameSkip
        
        'Frame Setup
        '--------------------------------
        FrameCNT = FrameCNT + 1
        If Not sectn = 1 Then
            Call FrameSetup(sectn, mfgpos)  '>>>>>>>>>>
            ReDim Preserve CollArray(1 To FrameCNT - 1)
            Set CollArray(FrameCNT - 1).CollEvents = Me.Controls(Section(sectn) & "Collapse" & mfgpos)
        End If
    
        'add Nav and Collapse ctrls to event handler
        '--------------------------------
        ReDim Preserve NavArray(1 To FrameCNT)
        Set NavArray(FrameCNT).NavEvents = Me.Controls(Section(sectn) & "Nav" & mfgpos)

        'add itms
        '--------------------------------
        For Ctrl = 1 To UBound(DataCtrl)
            If sectn = 1 Then
                CtrlCol = MemberValCol(Ctrl)
                For itm = 1 To mtchcats
                    Call AddDataCtrl(sectn, Ctrl, mfgpos, itm, CtrlCol)  '>>>>>>>>>>
                Next
            ElseIf sectn = 2 Then
                CtrlCol = MbrBenchValCol(Ctrl)
                If Ctrl = 2 Or Ctrl = 4 Or Ctrl = 6 Then
                    Call AddDataCtrl(sectn, Ctrl, mfgpos, 1, CtrlCol)  '>>>>>>>>>>
                ElseIf Ctrl = 8 Then
                    For itm = 1 To mtchcats
                        Call AddDataCtrl(sectn, Ctrl, mfgpos, itm, CtrlCol)  '>>>>>>>>>>
                    Next
                End If
            ElseIf sectn = 3 Then
                On Error Resume Next
                CtrlCol = Left(Range(SupplierValCol(Ctrl) & 1).Offset(0, (mfgpos - 1) * 30).Address(0, 0), 2)
                On Error GoTo 0
                If Ctrl = 2 Or Ctrl = 3 Or Ctrl = 4 Or Ctrl = 6 Then
                    Call AddDataCtrl(sectn, Ctrl, mfgpos, 1, CtrlCol)  '>>>>>>>>>>
                ElseIf Ctrl = 8 Then
                    For itm = 1 To mtchcats
                        Call AddDataCtrl(sectn, Ctrl, mfgpos, itm, CtrlCol)  '>>>>>>>>>>
                    Next
                End If
            ElseIf sectn = 4 Then
                On Error Resume Next
                CtrlCol = Left(Range(SuppBenchValCol(Ctrl) & 1).Offset(0, (mfgpos - 1) * 30).Address(0, 0), 2)
                On Error GoTo 0
                If Ctrl = 2 Or Ctrl = 4 Or Ctrl = 6 Then
                    Call AddDataCtrl(sectn, Ctrl, mfgpos, 1, CtrlCol)  '>>>>>>>>>>
                ElseIf Ctrl = 8 Then
                    For itm = 1 To mtchcats
                        Call AddDataCtrl(sectn, Ctrl, mfgpos, itm, CtrlCol) '>>>>>>>>>>
                    Next
                End If
            End If
        Next
        
        'add UOM ctrls to event handler
        '--------------------------------
        If sectn = 1 Then
            For itm = 1 To mtchcats
                UOMtxtCNT = UOMtxtCNT + 1
                ReDim Preserve UOMtxtArray(1 To UOMtxtCNT)
                Set UOMtxtArray(UOMtxtCNT).UOMEvents = Me.Controls(Section(1) & DataCtrl(2) & 1 & itm)
            Next
        Else
            UOMtxtCNT = UOMtxtCNT + 1
            ReDim Preserve UOMtxtArray(1 To UOMtxtCNT)
            Set UOMtxtArray(UOMtxtCNT).UOMEvents = Me.Controls(Section(sectn) & DataCtrl(2) & mfgpos & 1)
        End If
        
        'set frame height
        '--------------------------------
        If Not sectn = 1 Then
            FormWdth = FormWdth + FrameWdth
            Me.Controls(Section(sectn) & "Frame" & mfgpos).Width = FrameWdth
            Me.Controls(Section(sectn) & "Collapse" & mfgpos).Left = FormWdth - Me.Controls(Section(sectn) & "Collapse" & mfgpos).Width
            Me.Controls(Section(sectn) & "Nav" & mfgpos).Width = FrameWdth
        End If
        Me.Controls(Section(sectn) & "Frame" & mfgpos).Height = FrameHght
 
FrameSkip:
    Next

SectionSkip:
Next

'Change color of contracted supplier nav
'------------------------------------
If Not ConCol = 0 Then
    Me.Controls(Section(3) & "Nav" & conpos).BackColor = &H80FF80
    Me.Controls(Section(3) & "Collapse" & conpos).BackColor = &H80FF80
End If
    
'Set HxW
'------------------------------------
Call setHgtWdth(FormHght, FormWdth)

'NxtBttn_Click(true) 'calculate how many items left
CIAFrame.Visible = False
AllCats.EntireRow.Select


Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::



End Sub
Sub AddDataCtrl(sectn, Ctrl, mfgpos, itm, CtrlCol)
                
         
    Set CtrlModel = Me.Controls(Section(sectn) & Replace(DataCtrl(Ctrl), "Data", "Hdr") & mfgpos)   '<<set format model from member section
    Set CurrFrame = Me.Controls(Section(sectn) & "Frame" & mfgpos)
    If Ctrl = 2 Then
        Set DataAdd = CurrFrame.Controls.Add("Forms.textbox.1", Section(sectn) & DataCtrl(Ctrl) & mfgpos & itm, True)
    ElseIf Ctrl = 7 Then
        Set DataAdd = CurrFrame.Controls.Add("Forms.togglebutton.1", Section(sectn) & DataCtrl(Ctrl) & mfgpos & itm, True)
    Else
        Set DataAdd = CurrFrame.Controls.Add("Forms.Label.1", Section(sectn) & DataCtrl(Ctrl) & mfgpos & itm, True)
    End If
    
    With DataAdd
        CtrlVal = Range(CtrlCol & FirstCat.Offset(itm - 1, 0).Row).Value
        CtrlSize = CtrlModel.Font.Size
        If Ctrl = 2 Then
            CtrlSize = CtrlModel.Font.Size - 1
            If sectn = 4 Then
                'dataAdd.Height = 0  '(must change to 0 so that UOMchangeEvent does not fire during runcat)
                .Text = Sheets("Best Market Price").Range("A:A").Find(what:=CtrlVal, lookat:=xlWhole).Offset(0, 10).Value
                'dataAdd.Height = arraymodel.Height
            ElseIf sectn = 2 Then
                .Text = Sheets("Best Market Price").Range("A:A").Find(what:=CtrlVal, lookat:=xlWhole).Offset(0, 10).Value
            Else
                .Text = CtrlVal
            End If
        ElseIf Ctrl = 7 Then
            If Trim(LCase(CtrlVal)) = "x" Then
                DataAdd.Picture = LoadPicture(IconsPATH & "\" & XoutIcon, CtrlModel.Width + 2, CtrlModel.Height + 2)
                'DataAdd.BackColor = &H0&
                DataAdd.Value = True
            End If
            ReDim Preserve Xarray(1 To itm)
            Set Xarray(itm).XEvents = DataAdd
        Else
            If Ctrl = 8 And CtrlVal = "" Then CtrlVal = "-"
            .Caption = CtrlVal
            If (Ctrl = 4 Or Ctrl = 5) And Val(DataAdd.Caption) = 0 Then DataAdd.BackColor = &HFF&
            If Ctrl = 4 Or Ctrl = 6 Then
                If Left(CtrlVal, 1) = 0 And Len(CtrlVal) > 4 And Not CtrlVal = 0 Then
                    DataAdd.Caption = Format(CtrlVal, "$0.0000")
                Else
                    DataAdd.Caption = Format(CtrlVal, "$0.00")
                End If
            End If
            
            'Format Variance color
            '------------------------
            If Ctrl = 8 Then
                If sectn = 1 Then
                    DataAdd.Left = DataAdd.Left + 6
                    If DataAdd.Caption > plvlVarRange Or DataAdd.Caption < -plvlVarRange Then DataAdd.BackColor = &HFF&
                ElseIf sectn = 2 And (DataAdd.Caption > suppVarRange Or DataAdd.Caption < -suppVarRange) Then
                    DataAdd.BackColor = &HFF&
                ElseIf DataAdd.Caption > BnchVarRange Or DataAdd.Caption < -BnchVarRange Then
                    DataAdd.BackColor = &HFF&
                End If
                DataAdd.Caption = Format(DataAdd.Caption, "0%")
                If CtrlVal = "-" Then
                    If LCase(Range("AI" & FirstCat.Offset(itm - 1, 0).Row).Value) = "x" Then
                        DataAdd.BackColor = &H8000000F
                    Else
                        DataAdd.BackColor = &HFF&
                    End If
                End If
            End If
        End If
        
        .Top = CtrlModel.Top + CtrlModel.Height * itm
        .Left = CtrlModel.Left
        .Font.Size = CtrlSize
        .Font.Name = CtrlModel.Font.Name
        '.Font.Bold = CtrlModel.Font.Bold
        .BackStyle = CtrlModel.BackStyle
        .TextAlign = CtrlModel.TextAlign
        .Height = CtrlModel.Height
        .Width = CtrlModel.Width
    End With





End Sub
Sub FrameSetup(sectn, mfgpos)


            'add section frame
            '------------------------
            Set CtrlModel = Me.Controls(Section(1) & "Frame" & 1)
            Set FrameAdd = Me.Controls.Add("forms.frame.1", Section(sectn) & "Frame" & mfgpos, True)
            With FrameAdd
                .Caption = vbNullString
                .BackColor = CtrlModel.BackColor
                .BorderStyle = CtrlModel.BorderStyle
                '.BackStyle = MemberFrame1.BackStyle
                '.Height = MemberFrame1.Height
                .Left = FormWdth
                .Top = CtrlModel.Top
            End With
'            If Sectn = 3 Then
'                SuppFrames = SuppFrames + 1
'            ElseIf Sectn = 4 Then
'                BenchFrames = BenchFrames + 1
'            End If
            
            'add Nav button
            '===============================================================================
            
            'Get nav name
            '------------------------
            If sectn = 2 Then
                NavName = "10% Benchmark"
                navcolor = &HFFFFC0
            ElseIf sectn = 3 Then
                NavName = ConTblBKMRK.Offset(0, mfgpos).Value
                navcolor = &HFFFF&
            ElseIf sectn = 4 Then
                NavName = ConTblBKMRK.Offset(0, mfgpos).Value
                navcolor = &HFF&
            End If
            
            'add nav ctrl
            '------------------------
            Set CtrlModel = Me.Controls(Section(1) & "Nav" & 1)
            Set NavCtrl = FrameAdd.Controls.Add("Forms.CommandButton.1", Section(sectn) & "Nav" & mfgpos, True)
            With NavCtrl
                .Caption = NavName
                .Font.Size = CtrlModel.Font.Size
                .Font.Name = CtrlModel.Font.Name
                .Font.Bold = CtrlModel.Font.Bold
                .BackColor = navcolor
                '.TextAlign = HomeUOMVal.TextAlign
                .Height = CtrlModel.Height
                .Left = 0 'FormWdth
                .Top = 0 'MemberNav1.Top
            End With
            
            'add collpase button
            '===============================================================================
            Set CtrlModel = Me.Controls(Section(1) & "Collapse" & 1)
            Set CollAdd = Me.Controls.Add("Forms.togglebutton.1", Section(sectn) & "Collapse" & mfgpos, True)
            With CollAdd
                .Caption = "~"
                .Font.Size = CtrlModel.Font.Size
                .Font.Name = CtrlModel.Font.Name
                .Font.Bold = CtrlModel.Font.Bold
                .BackColor = navcolor
                .BackStyle = CtrlModel.BackStyle
                .TextAlign = CtrlModel.TextAlign
                .Height = CtrlModel.Height
                .Width = CtrlModel.Width
                .Top = FrameAdd.Top - CtrlModel.Height
                .Picture = LoadPicture(IconsPATH & "\" & CollapseIcon, CollAdd.Width, CollAdd.Height)
            End With
            
            'add fields to frame
            '===========================================================================================================================================
            FrameWdth = 0
            For Ctrl = 1 To 8
                
                Set CtrlModel = Me.Controls(Section(1) & Replace(DataCtrl(Ctrl), "Data", "Hdr") & 1)   '<<set format model from member section

                If Ctrl = 1 Or Ctrl = 5 Or Ctrl = 7 Then GoTo CtrlSkip   '<<skip orgnl, x, usage for supp & bench sections
                If Not sectn = 3 And Ctrl = 3 Then GoTo CtrlSkip         '<<skip Pkg for Bench and Ten Sections
                
                'add header
                '------------------------
                Set HdrAdd = FrameAdd.Controls.Add("Forms.Label.1", Section(sectn) & Replace(DataCtrl(Ctrl), "Data", "Hdr") & mfgpos, True)
                With HdrAdd
                    .Caption = CtrlModel.Caption
                    .Font.Size = CtrlModel.Font.Size
                    .Font.Name = CtrlModel.Font.Name
                    .Font.Bold = CtrlModel.Font.Bold
                    .BackStyle = CtrlModel.BackStyle
                    .TextAlign = CtrlModel.TextAlign
                    .Width = CtrlModel.Width
                    .Height = CtrlModel.Height
                    .Left = FrameWdth
                    .Top = CtrlModel.Top
                End With
                FrameWdth = FrameWdth + CtrlModel.Width
CtrlSkip:
            Next
                
End Sub
Sub setHgtWdth(HeightVar As Double, WidthVar As Double)


'set form height x width & scrollbars
'=============================================================================================================
If HeightVar > Application.Height - 50 Then
    With Me
        .Height = Application.Height - 30
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = RubiksForm.Height + Abs(HeightVar - RubiksForm.Height) '- 30) ' 'Me.Height '2 '.InsideHeight * 2
    End With
Else
    If mtchcats = 1 Then
        Me.Height = 102
    ElseIf mtchcats = 2 Then
        Me.Height = (MemberFrame1.Height + CatTitle.Height) * 1.21
    ElseIf mtchcats < 5 Then
        Me.Height = (MemberFrame1.Height + CatTitle.Height) * 1.17
    ElseIf mtchcats < 10 Then
        Me.Height = (MemberFrame1.Height + CatTitle.Height) * 1.15
    ElseIf mtchcats < 20 Then
        Me.Height = (MemberFrame1.Height + CatTitle.Height) * 1.11
    ElseIf mtchcats < 30 Then
        Me.Height = (MemberFrame1.Height + CatTitle.Height) * 1.07
    Else
        Me.Height = (MemberFrame1.Height + CatTitle.Height) * 1.05
    End If
    Me.ScrollBars = 0
End If

If WidthVar > Application.Width - 5 Then
    With Me
        .Width = Application.Width - 20
        .ScrollBars = fmScrollBarsHorizontal
        .ScrollWidth = RubiksForm.Width + Abs(WidthVar - RubiksForm.Width) 'WidthVar '2 '.InsideHeight * 2
        .Height = Me.Height + 10
    End With
ElseIf WidthVar < CIAbttn.Left + CIAbttn.Width Then
    Me.Width = CIAbttn.Left + CIAbttn.Width + 5
Else
    Me.Width = WidthVar + 5
    Me.ScrollBars = 0
End If


End Sub
Private Sub OrgnlValues_Click()

''check to make sure catalog number in col N is activecell
''---------------------------------
'Set isect = Application.Intersect(ActiveCell, Range("N:N"))
'If isect Is Nothing Then
'    MsgBox "Please select the catalog number you wish to evaluate in column N."
'    Exit Sub
'End If

Sheets("line Item data").Select

'set catvals
'---------------------------------
'Catnmbr = ActiveCell.Offset(0, Range("X1").Column - ActiveCell.Column).Value
'Set FirstCat = Range("X:X").Find(what:=Catnmbr, lookat:=xlWhole)
'mtchcats = Application.CountIf(Range("X:X"), FirstCat.Value)                  'Count number of matching catalog numbers
'Set AllCats = Range(FirstCat, FirstCat.Offset(mtchcats - 1, 0))

'check to make sure sorted by catalog number
'---------------------------------
For Each c In AllCats
    If c.Value = FirstCat.Value Then
        contigchk = contigchk + 1
    End If
Next
If Not contigchk = mtchcats Then
    MsgBox "Please make sure your sheet is sorted by catalog number and try again."
    Exit Sub
End If

'Reset mbr UOMs
'---------------------------------
For Each c In AllCats
    c.Offset(0, 4).Value = c.Offset(0, 1).Value
Next

'reset mbr bench UOM
'---------------------------------
If Not FirstCat.Offset(0, 20).Value = "-" Then
    Set MbrBenchCat = Sheets("Best Market Price").Range("A:A").Find(what:=Catnmbr, lookat:=xlWhole)
    MbrBenchCat.Offset(0, 10).Value = MbrBenchCat.Offset(0, 15).Value
    MbrBenchCat.EntireRow.Calculate
End If

'reset supp And supp Bench UOMs
'---------------------------------
On Error GoTo EndClean
mfgcol = Rows("4:4").Find(what:=FirstCat.Offset(0, -2), lookat:=xlWhole).Column
suppcatnmbr = FirstCat.Offset(0, mfgcol - FirstCat.Column).Value
If Not suppcatnmbr = "-" Then
    Set suppcat = Sheets(FUN_SuppName & " pricing").Range("A:A").Find(what:=suppcatnmbr, lookat:=xlWhole)
    suppcat.Offset(0, 4).Value = Sheets(FUN_SuppName & " pricing").Rows("1:1").Find(what:="Original UOMs").Offset(suppcat.Row - 1, 0).Value
    suppcat.EntireRow.Calculate
    
    Set SuppBenchCat = Sheets("Best Market Price").Range("A:A").Find(what:=suppcatnmbr, lookat:=xlWhole)
    SuppBenchCat.Offset(0, 10).Value = SuppBenchCat.Offset(0, 15).Value
    SuppBenchCat.EntireRow.Calculate
End If

EndClean:
AllCats.Offset(0, 11).ClearContents
AllCats.Offset(0, 11).Interior.ColorIndex = 0
AllCats.Offset(0, 10).Interior.ColorIndex = 0
AllCats.EntireRow.Calculate


End Sub
Sub AcceptBttn_click()

'check sort
'---------------------------------
endFLG = 0
Call SortChk
If endFLG = 1 Then
    endFLG = 0
    Exit Sub
End If

'input mbr values
'---------------------------------
For i = 1 To mtchcats
    If Not FirstCat.Offset(i - 1, 4).Value = Val(Me.Controls(Section(1) & DataCtrl(2) & 1 & i).Text) Then FirstCat.Offset(i - 1, 4).Value = Me.Controls(Section(1) & DataCtrl(2) & 1 & i).Text
    If Me.Controls(Section(1) & DataCtrl(7) & 1 & i).Value = True Then
        FirstCat.Offset(i - 1, 11).Value = "x"
    Else
        FirstCat.Offset(i - 1, 11).Value = ""
    End If
Next

'Input new member benchmark values
'---------------------------------
If Not Sheets("Line Item Data").Range("AR" & FirstCat.Row).Value = "-" Then
    Set MbrBenchUOM = Sheets("Best Market Price").Range("A:A").Find(what:=Catnmbr, lookat:=xlWhole).Offset(0, 10)
    If Not MbrBenchUOM.Value = Val(Me.Controls(Section(2) & DataCtrl(2) & 11).Text) Then
        MbrBenchUOM.Value = Val(Me.Controls(Section(2) & DataCtrl(2) & 11).Text)
        MbrBenchUOM.EntireRow.Calculate
    End If
End If

For i = 1 To suppNMBR
    
    'input supplier values
    '---------------------------------
    If Not Sheets("Line Item Data").Range("BG" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "-" Then
        Set suppUOM = Sheets("Line Item Data").Range("BI" & FirstCat.Row).Offset(0, (i - 1) * 30)
        If Not suppUOM.Value = Val(Me.Controls(Section(3) & DataCtrl(2) & i & 1).Text) Then
            Set PricefileUOM = Sheets(FUN_SuppName(i) & " Pricing").Range("A:A").Find(what:=suppUOM.Offset(0, -2).Value, lookat:=xlWhole).Offset(0, 4)
            PricefileUOM.Value = Val(Me.Controls(Section(3) & DataCtrl(2) & i & 1).Text)
            PricefileUOM.EntireRow.Calculate
        End If
    End If
    
    'input supplier Benchmark values
    '---------------------------------
    If Not Sheets("Line Item Data").Range("BV" & FirstCat.Row).Offset(0, (i - 1) * 30).Value = "" Then
        BenchCatNmbr = Sheets("Line Item Data").Range("BG" & FirstCat.Row).Offset(0, (i - 1) * 30).Value
        Set SuppBenchUOM = Sheets("Best Market Price").Range("A:A").Find(what:=BenchCatNmbr, lookat:=xlWhole).Offset(0, 10) 'Range(SuppValArray(2) & FirstCat.Row).Offset(0, i * 18)
        If Not SuppBenchUOM.Value = Val(Me.Controls(Section(4) & DataCtrl(2) & i & 1).Text) Then
            SuppBenchUOM.Value = Val(Me.Controls(Section(4) & DataCtrl(2) & i & 1).Text)
            SuppBenchUOM.EntireRow.Calculate
        End If
    End If
    
Next

Call Calculate_Priceleveling_Single(AllCats.Offset(0, 15))
AllCats.EntireRow.Calculate



End Sub
Sub NxtBttn_Click()

    Sheets("Line Item Data").Select
    
    '(in case user clicks the next button before starting with the run button)
    If FirstCat = "" Then
        Set FirstCat = Range("X4")
        mtchcats = 1
    End If
    
    If FirstCat.Offset(mtchcats, 0).Row > ItmNmbr + 4 Then
        MsgBox "All variances have been evaluated"
        Exit Sub
    End If
    
    If menuViewAll = True Then
        FirstCat.Offset(mtchcats, 0).Select
        Call RunCat_Click
        Exit Sub
    End If
    
    For Each itm In Range(FirstCat.Offset(mtchcats, 0), Range("X" & ItmNmbr + 4))
        
        'Check if already evaluated
        '-------------------------------------
        If itm.Offset(0, 3).Value = itm.Offset(0, 4).Value And Not LCase(itm.Offset(0, 11).Value) = "x" Then
                
            'if UOMqty is 1 and UOMpkg is not an EA
            '-------------------------------------
            If itm.Offset(0, 4).Value = 1 And (itm.Offset(0, 5).Value = "BN" Or itm.Offset(0, 5).Value = "CA" Or itm.Offset(0, 5).Value = "PL" Or itm.Offset(0, 5).Value = "BG" Or itm.Offset(0, 5).Value = "BX" Or itm.Offset(0, 5).Value = "CT" Or itm.Offset(0, 5).Value = "DZ" Or itm.Offset(0, 5).Value = "PK") Then
                
                'check if more mismatchs than mismatch threshold [default = 4/6 (65%)]
                '-------------------------------------
                Set itmFirstCat = Range("X:X").Find(what:=itm.Value, lookat:=xlWhole)
                itmMtchcats = Application.CountIf(Range("X:X"), itmFirstCat.Value)
                If Application.CountIf(Range(itmFirstCat.Offset(0, 5), itmFirstCat.Offset(itmMtchcats - 1, 5)), "EA") / itmMtchcats < Val(Format(ZeusForm.msmtchSelect.Value, "0.00")) Then GoTo NxtRun
            
            End If
                
            'check for 0s in unit cost and usage
            '-------------------------------------
            If itm.Offset(0, 6) = 0 Or itm.Offset(0, 9) = 0 Then GoTo NxtRun '()
            
            'check price leveling variance
            '-------------------------------------
            If Range("AO" & itm.Row).Value > plvlVarRange Or Range("AO" & itm.Row).Value < -plvlVarRange Then GoTo NxtRun
            
            'check member benchmark variance
            '-------------------------------------
            If Not Range("AR" & itm.Row).Value = "-" And (Range("AV" & itm.Row).Value > BnchVarRange Or Range("AV" & itm.Row).Value < -BnchVarRange) Then GoTo NxtRun
            
            'check supplier variance and supplier benchmark variance
            '-------------------------------------
            For i = 1 To suppNMBR
                If Not Range("BG" & itm.Row).Offset(0, (i - 1) * 30).Value = "-" And (Range("BN" & itm.Row).Offset(0, (i - 1) * 30).Value > suppVarRange Or Range("BN" & itm.Row).Offset(0, (i - 1) * 30).Value < -suppVarRange) Then GoTo NxtRun
                If Not Trim(Range("BV" & itm.Row).Offset(0, (i - 1) * 30).Value) = "" And (Range("BZ" & itm.Row).Offset(0, (i - 1) * 30).Value > BnchVarRange Or Range("BZ" & itm.Row).Offset(0, (i - 1) * 30).Value < -BnchVarRange) Then GoTo NxtRun
            Next
            
        End If
        GoTo NxtItm
NxtRun:
'        If ItemsLeft = True Then
'            itmCnt = itmCnt + 1
'        Else
            itm.Select
            Call RunCat_Click
            Exit Sub
'        End If
NxtItm:
    Next


MsgBox "All variances have been evaluated"
Exit Sub

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::




End Sub
Sub PrevBttn_Click()

Sheets("Line Item Data").Select
If PrevCat = "" Then Exit Sub
PrevCat.Select
Call RunCat_Click


End Sub
Sub googlebttn_click()


Dim hwnd As Long, IECaption As String
Dim HTMLDoc As HTMLDocument
Dim cardinalHTML As HTMLDocument
Dim oHTML_Element As IHTMLElement
Dim sURL As String

'ttlwait3 = 0
'ttlwait4 = 0
'ttlwait5 = 0
'ttlwait6 = 0
'Shell ("C:\Users\USERNAME\AppData\Local\Google\Chrome\Application\Chrome.exe -url http:google.ca")

'open browser
'===============================================================================
'On Error GoTo Err_Clear
mfg = FirstCat.Offset(0, -2).Value
sURL = "https://www.google.com/?gws_rd=ssl#safe=strict&q=" & Catnmbr & "+" & mfg
'sURL = "https://mysite.cardinal.com/CAH/mpsGuest.jsp"

On Error GoTo errhnldNoBrowser
NoBrowser:
If RubiksBrowser Is Nothing Then

    Set RubiksBrowser = New InternetExplorer
    
    'if excel is on left monitor and not on laptop then open right (Assumes laptop resolution is not more than 1400pixels and that if >1400 pixels then user is not only using one monitor and that user has excel maximized and not hanging off the edge of the screen)
    '-------------------------------------------------------
    If Application.Left < 1400 Then
        RubiksBrowser.Left = Application.Left + Application.Width + 1000
        'Debug.Print "1st"
    Else
        RubiksBrowser.Left = 0
        'Debug.Print "second"
    End If
    
    RubiksBrowser.Navigate sURL
    sURL = "https://mysite.cardinal.com/CAH/mpsGuest.jsp"
    RubiksBrowser.Navigate sURL, 2048&
    
'    While RubiksBrowser.Busy Or RubiksBrowser.ReadyState <> READYSTATE_COMPLETE: DoEvents: Wend
'    Set cardinalHTML = RubiksBrowser.Document
'
'LoadCheck:
'    For Each elem In cardinalHTML.getElementsByTagName("input")
'        'If elem.ID = "ucSearchBasic_txtSearchText" Then
'        Debug.Print elem.outerHTML
'        If InStr(elem.outerHTML, "ucSearchBasic_txtSearchText") > 0 Then
'            loadedFlg = True
'            Exit For
'        End If
'    Next
'    If Not loadedFlg = True Then GoTo LoadCheck
'    cardinalHTML.getElementsByID("ucSearchBasic_txtSearchText").Value = FirstCat.Value
Else
    RubiksBrowser.Navigate sURL
End If

RubiksBrowser.Visible = True
apiShowWindow RubiksBrowser.hwnd, SW_SHOWMAXIMIZED

On Error GoTo 0

'For Each elem In cardinalHTML.getElementsByTagName("frame")
'    If elem.Name = "basefrm" Then Set frameobj = elem
'Next
'
'For Each elm In frameobj.Document.getElementsByTagName("meta")
'    Debug.Print elm.innerHTML
'Next
'
'Set tstfrm = cardinalHTML.getElementsByTagName("frame").Name("basefrm")
'For i = 0 To cardinalHTML.frames
'
'    Debug.Print IE.Document.frames(i).Name
'    'Sheet1.Cells(i + 1, "C").Value = IE.Document.frames(i).Location
'Next
'Set HTMLDoc = cardinalHTML.frames("basefrm")
'HTMLDoc.DocumentElement.outerHTML

'While RubiksBrowser.Busy Or RubiksBrowser.READYSTATE <> READYSTATE_COMPLETE: DoEvents: Wend
'
'Debug.Print HTMLDoc.body
'
'RubiksBrowser.Document.getElementById("ucSearchBasic_txtSearchText").Value = Catnmbr
'HTMLDoc.getElementById("ucSearchBasic_btnSearch").Click
'
'
'   For Each oHTML_Element In HTMLDoc.getElementsByTagName("select")
'    Debug.Print oHTML_Element.Name
'    If oHTML_Element.Name = "selectCatalog" Then
'        'oHTML_Element.Value = "Full Products Catalog"
'        oHTML_Element.selectedIndex = 1
'        oHTML_Element.FireEvent ("onchange")
'
'
'        Exit For
'    End If
'Next
'
'    HTMLDoc.all.ucSearchBasic_btnSearch.Click
'For Each oHTML_Element In HTMLDoc.getElementsByID("input")
'    Debug.Print oHTML_Element.Type
'    If oHTML_Element.ID = "ucSearchBasic_txtSearchText" Then
'        oHTML_Element.Click
'        Exit For
'    End If
'Next
'
'
'    While RubiksBrowser.Busy Or RubiksBrowser.READYSTATE <> READYSTATE_COMPLETE: DoEvents: Wend

Exit Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
errhnldNoBrowser:
Set RubiksBrowser = Nothing
On Error GoTo 0
Resume NoBrowser

End Sub
Sub CIAbttn_click()

'resize
'-------------------------
CIAFrame.Visible = True
CIAFrame.Left = 0
CIAFrame.Top = MemberFrame1.Top + MemberFrame1.Height
CIAFrame.Width = MemberFrame1.Width


For Each clrframe In CIAFrame.Controls
    If InStr(LCase(clrframe.Name), "data") Then Me.Controls.Remove clrframe.Name
Next

Sheets("Line Item Data").Select

If Not RunCatFLG = 1 Then

    'lookup UOMs
    '---------------------------
    ProdNmbr = Catnmbr
    colnmbr = Application.CountA(Sheets("Line Item Data").Rows("4:4"))
    Set UOMref = Sheets("Line Item Data").Range("A4").Offset(0, colnmbr)
    'Range(UOMref, UOMref.End(xlToRight)).EntireColumn.ClearContents
    UOMref.Value = "x"   'placeholder for einstein notes
    Range(UOMref, UOMref.Offset(0, -1)).Columns.Hidden = False

    Set ReturnCol = Sheets("Line Item Data").Range("A4").End(xlToRight).Offset(0, 10)
    ReturnCol.EntireColumn.ClearContents
    Range(ReturnCol.Offset(-3, 0), ReturnCol).Value = "x"
    Call UOMcrawler_CIA(ProdNmbr, ReturnCol)  '>>>>>>>>>>

End If

'cross with contextUOMs and add to UOMset if good refs
'---------------------------
CIAhght = CIAnav.Top + CIAnav.Height '.Top + CIAUOM.Height
If UOMfndFLG = 1 Then
'    rownmbr = 1
'    UOMnmbr = 1
    For Each c In Range(UOMref.Offset(0, 1), UOMref.End(xlToRight))
        If Not InStr(c.Value, "Desc") > 0 Then
            
            'resize UOM label width
            '-------------------------
            'If IsNumeric(Left(Trim(c.Value), 1)) Then
            initlen = Len(c.Value)
            SlashCnt = initlen - Len(Replace(c.Value, "/", ""))
            labelwdth = SlashCnt * 42
        
            'add itm to frame
            '-------------------------
            UOMnmbr = UOMnmbr + 1
            rownmbr = rownmbr + 1
            Set ArrayModel = Me.CIAUOM
            Set DataAdd = CIAFrame.Controls.Add("Forms.Label.1", "UOMdata" & rownmbr & UOMnmbr, True)
            With DataAdd
                .Caption = c.Value
                .Font.Size = ArrayModel.Font.Size
                .TextAlign = 1
                .Width = labelwdth
                .Height = ArrayModel.Height
                .Left = UOMleft
                .Top = CIAhght
            End With
            UOMleft = UOMleft + DataAdd.Width
            If UOMleft > Maxleft Then Maxleft = UOMleft
        Else
            'add to desc and go to next row
            '--------------
            Set ArrayModel = Me.CIAUOM
            Set DataAdd = CIAFrame.Controls.Add("Forms.textbox.1", "Descdata" & rownmbr, True)
            With DataAdd
                .Text = c.Value
                .Font.Size = ArrayModel.Font.Size
                .TextAlign = 1
                .Width = CIAnav.Width
                .Height = ArrayModel.Height * 1.5
                .Left = 2
                .Top = CIAhght + DataAdd.Height
            End With
            CIAhght = CIAhght + DataAdd.Height * 2
            UOMleft = 0
        End If
        
    Next
Else
    Set ArrayModel = Me.CIAUOM
    Set DataAdd = CIAFrame.Controls.Add("Forms.Label.1", "Descdata1", True)
    With DataAdd
        .Caption = "No Data Available"
        .Font.Size = ArrayModel.Font.Size
        .TextAlign = ArrayModel.TextAlign
        .Width = 50
        .Height = ArrayModel.Height
        .Left = 2
        .Top = CIAhght
    End With
    CIAhght = CIAhght + DataAdd.Height
    Maxleft = 0
End If

Range(UOMref, UOMref.End(xlToRight)).ClearContents
ReturnCol.EntireColumn.ClearContents


'Adjust HxW
'---------------------
CIAFrame.Height = CIAhght + 3
If rownmbr = "" Then
    RubiksForm.Height = CIAFrame.Top + CIAFrame.Height * 1.53
ElseIf rownmbr = 2 Then
    RubiksForm.Height = CIAFrame.Top + CIAFrame.Height * 1.21  'CIAhght * 1.5
Else
    RubiksForm.Height = CIAFrame.Top + CIAFrame.Height * 1.17
End If
If Maxleft > CIAnav.Width Then
    CIAnav.Width = Maxleft
    CIAFrame.Width = Maxleft
End If
For Each c In CIAFrame.Controls
    If InStr(LCase(c.Name), "descdata") Then c.Width = CIAnav.Width
Next
If CIAFrame.Width > FormWdth Then FormWdth = CIAFrame.Width


'Call setHgtWdth(CIAFrame.Top + CIAFrame.Height, FormWdth)



End Sub
Sub HomeMenu_click()


If MenuFrame.Visible = False Then
    MenuFrame.Visible = True
    MenuFrame.ZOrder 0  '
    'MemberFrame1.ZOrder 1
    With MenuFrame
        .Left = HomeMenu.Left + HomeMenu.Width
        .Top = HomeMenu.Top
    End With
    If Not Me.Height > MenuFrame.Height + 22 Then
        PrevHeight = Me.Height
        Me.Height = MenuFrame.Height + 22
    End If
'    MemberFrame1.Visible = False
'    MemberFrame1.Visible = True
'    If CIAFrame.Visible = True Then
'        CIAFrame.Visible = False
'        CIAFrame.Visible = True
'    End If
Else
    If Me.Height = MenuFrame.Height + 22 Then Me.Height = PrevHeight
    MenuFrame.Visible = False
    AddUOMframe.Visible = False
    TargEAFrame.Visible = False
End If


End Sub
Private Sub MenuFrame_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    'Adjust HxW
    '-----------------
    If Me.Height = MenuFrame.Height + 22 Then Me.Height = PrevHeight

    MenuFrame.Visible = False
    AddUOMframe.Visible = False
    TargEAFrame.Visible = False

End Sub
Sub addHansNote_Click()

'add CIA data if already looked up
'--------------------------------------------------
If CIAFrame.Visible = True Then
    WebRefNote = "("
    For Each Ctrl In RubiksForm.CIAFrame.Controls
        If InStr(Ctrl.Name, "UOMdata") Then WebRefNote = WebRefNote & Ctrl.Caption & "; "
    Next
    If WebRefNote = "(" Then
        WebRefNote = vbNullString
    Else
        WebRefNote = WebRefNote & "CIA)"
    End If
Else
    WebRefNote = vbNullString
End If

'show notes form
'--------------------------------------------------
AddNoteForm.Show (False)
AddNoteForm.Left = RubiksForm.AddHansNote.Left + RubiksForm.AddHansNote.Width
AddNoteForm.Top = RubiksForm.AddHansNote.Top / 2

'Adjust HxW
'--------------------------------------------------
If Me.Height = MenuFrame.Height + 22 Then Me.Height = PrevHeight
MenuFrame.Visible = False
AddUOMframe.Visible = False
TargEAFrame.Visible = False




End Sub
'Private Sub ShowZeus_Click()
'
'If ShowZeus = False Then
'    ZeusForm.Show (False)
'    ZeusForm.Left = RubiksForm.Left
'    If Application.Top > RubiksForm.Top - ZeusForm.Height Then   'if top is relative to application then>> if RubiksForm.Top - zeusform.Height < 0 then
'        ZeusForm.Top = RubiksForm.Top - ZeusForm.Height
'    Else
'        ZeusForm.Top = RubiksForm.Top
'    End If
'Else
'    ZeusForm.Hide
'End If
'
'
'End Sub
Private Sub RubiksForm_Deactivate()

'If MenuFrame.Visible = True Then
'    If Me.Height = MenuFrame.Height + 22 Then
'        Me.Height = prevHeight
'    End If
'    MenuFrame.Visible = False
'    AddUOMframe.Visible = False
'    TargEAFrame.Visible = False
'End If

End Sub
Sub UserForm_Click()

'hide menu frame
'------------------
If MenuFrame.Visible = True Then
    If Me.Height = MenuFrame.Height + 22 Then Me.Height = PrevHeight
    MenuFrame.Visible = False
    AddUOMframe.Visible = False
    TargEAFrame.Visible = False
End If


End Sub
Private Sub calcbttn_Click()

On Error Resume Next
If AllCats.EntireRow.Address = Selection.EntireRow.Address Then
    AllCats.EntireRow.Calculate
Else
    AllCats.EntireRow.Calculate
    Selection.Calculate
End If
On Error GoTo 0

End Sub
Private Sub userform_terminate()

Application.ScreenUpdating = True
HansFLG = 0

'Turn calculations back on
'-------------------------
Call FUN_CalcBackOn


End Sub
