VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SherlockForm 
   Caption         =   "Clinical"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   OleObjectBlob   =   "SherlockForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SherlockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

AddToForm MIN_BOX

End Sub
Sub addSherlockNote_Click()

If LCase(ActiveSheet.Name) = "Line Item Data" Then
    Catnmbr = ActiveCell.Offset(0, Range("N1").Column - ActiveCell.Column).Value
ElseIf ActiveSheet.Name = "ClinicalQC" Then
    Catnmbr = ActiveCell.Offset(0, Range("G1").Column - ActiveCell.Column).Value
Else
    MsgBox "please select a catalog number on either the Clinical QC tab or the HCO detail tab."
    Exit Sub
End If

AddNoteForm.Show (False)

''[TBD]determine if on other screen for positioning
'If Application.Left < 1400 Then
'    addwidth = Application.Left + Application.Width + 1000
'Else
'    addwidth = 0
'End If
AddNoteForm.Left = SherlockForm.AddSherlockNote.Left + SherlockForm.AddSherlockNote.Width
AddNoteForm.Top = SherlockForm.AddSherlockNote.Top / 2


End Sub

Sub okbutton_Click()
    
    Application.Calculation = xlCalculationManual       'vvvvv
    Application.ScreenUpdating = False                  'vvvvv
    
    Sheets("Line Item Data").Select
    If Range("ZA2").Value <> 1 Then
        
        Range("ZA:ZJ,ZL:AAA").Clear
        If Worksheets("Line Item Data").FilterMode = True Then
            Rows("2:2").AutoFilter
        End If
        
        'count how many left for evaluation and mark if already colored
        '-------------------------------------------
        ttlleft = 0
        colorcnt = 0
        Do
            colorcnt = colorcnt + 1
            DoEvents
            Application.StatusBar = "Finding matching criteria: rows  " & colorcnt
            
            If Range("ZK2").Offset(colorcnt, 0).Value <> 1 Then
                'If colorfunction(Range("A1"), Range("C2,E2,H2,L2,N2").Offset(colorcnt, 0), False) = 0 Then
                If colorfunction(Range("A1"), Range("H2,L2,N2").Offset(colorcnt, 0), False) = 0 Then
                    ttlleft = ttlleft + 1
                Else
                    Range("ZK2").Offset(colorcnt, 0).Value = 1
                End If
            End If
        Loop Until IsEmpty(Range("A2").Offset(colorcnt, 0))
        
        ItemsLeft.Caption = "Items left for evaluation: " & ttlleft - 1
        
    End If
    
    Range("ZA2").Value = ""
   
    Call clinicalrun  '>>>>>>>>>>>

Application.StatusBar = False

End Sub
Private Sub UserForm_Initialize()
    
    'set common variables
    '=============================================================================================================
    SetupSwitch = FUN_SetupSwitch

    Application.Calculation = xlCalculationManual       'vvvvv
    Application.ScreenUpdating = False                  'vvvvv
    
    Sheets("Line Item Data").Select
    Range("ZA3:BBA100000").ClearContents
    
    If Worksheets("Line Item Data").FilterMode = True Then
        Rows("2:2").AutoFilter
    End If
    
    'count how many left for evaluation and mark if already colored
    '-------------------------------------------
    Range("ZA2").Value = 1
    ttlleft = 0
    colorcnt = 0
    Do
        colorcnt = colorcnt + 1
        DoEvents
        Application.StatusBar = "Starting Sherlock...Rows " & colorcnt
        
        'If colorfunction(Range("A1"), Range("C2,E2,H2,L2,N2").Offset(colorcnt, 0), False) = 0 Then
        If colorfunction(Range("A1"), Range("H2,L2,N2").Offset(colorcnt, 0), False) = 0 Then
            ttlleft = ttlleft + 1
        Else
            Range("ZK2").Offset(colorcnt, 0).Value = 1
        End If
 
    Loop Until IsEmpty(Range("A2").Offset(colorcnt, 0))
    
    ItemsLeft.Caption = "Items left for evaluation: " & ttlleft - 1

Application.StatusBar = False

SherlockFLG = 1
'[TBD]move zeusform to where sherlockform starts up

'AddToForm MIN_BOX

End Sub
Private Sub ViewAll_Click()

    If Range("ZA2").Value <> 1 Then
        
        Application.Calculation = xlCalculationManual       'vvvvv
        Application.ScreenUpdating = False                  'vvvvv
       
        Sheets("Line Item Data").Select
        Range("ZA3:ZJ100000").ClearContents
        Range("ZL:AAA").Clear
        
        If Worksheets("Line Item Data").FilterMode = True Then
            Rows("2:2").AutoFilter
        End If
       
       'count how many left for evaluation and mark if already colored
       '-------------------------------------------
       ttlleft = 0
       colorcnt = 0
       Do
           colorcnt = colorcnt + 1
           DoEvents
           Application.StatusBar = "Finding matching criteria: Rows " & colorcnt
           
           If Range("ZK2").Offset(colorcnt, 0).Value <> 1 Then
                'If colorfunction(Range("A1"), Range("C2,E2,H2,L2,N2").Offset(colorcnt, 0), False) = 0 Then
                If colorfunction(Range("A1"), Range("H2,L2,N2").Offset(colorcnt, 0), False) = 0 Then
                    ttlleft = ttlleft + 1
                Else
                    Range("ZK2").Offset(colorcnt, 0).Value = 1
                End If
           End If
           
       Loop Until IsEmpty(Range("A2").Offset(colorcnt, 0))
       
       ItemsLeft.Caption = "Items left for evaluation: " & ttlleft - 1
    End If
    Range("ZA2").Value = ""

SherlockForm.Hide
SherlockForm.Show False

    incallcnt = -1
    Do
        incallcnt = incallcnt + 1
        If Range("ZK3").Offset(incallcnt, 0).Value <> 1 Then
            Range("ZL3").Offset(incallcnt, 0).Value = 1
        End If
    Loop Until IsEmpty(Range("A3").Offset(incallcnt, 0))

Range("ZK1").Value = 1
Call clinicalrun  '>>>>>>>>>>

Application.StatusBar = False

End Sub
Private Sub userform_terminate()

Unload SherlockForm
Application.ScreenUpdating = False                  'vvvvv
'Sheets("Line Item Data").Range("ZA:ZJ,ZL:AAA").ClearContents
Sheets("Line Item Data").Range("ZA:AAA").ClearContents

End Sub
Private Sub closebutton_click()

Call userform_terminate '>>>>>

End Sub
Function colorfunction(rcolor As Range, rrange As Range, Optional sum As Boolean)

Dim rcell As Range
Dim lcol As Long
Dim vresult

lcol = rcolor.Interior.Color           '<--if Rcolor is set to a certain range then the color will be evaluated based on the color of that range, otherwise if it is left blank the color will default to no fill
If sum = True Then                          '<--If sum is set to true then it evaluates differently other than each cell
    For Each rcell In rrange
        If rcell.Interior.Color = lcol Then
            vresult = WorksheetFunction.sum(rcell, vresult)
        End If
    Next rcell
Else
    If Not LCase(Usr) = "mlemay" Then
        For Each rcell In rrange
            'Debug.Print rcell.Address
            'If rcell.Interior.ColorIndex = lcol Then
            If rcell.Interior.Color = RGB(0, 127, 0) Or rcell.Interior.Color = 65280 Or rcell.Interior.Color = RGB(180, 0, 0) Or rcell.Interior.Color = 652804 Then
                vresult = 1 + vresult
            End If
        Next rcell
    Else
        For Each rcell In rrange
            'Debug.Print rcell.Address
            'If rcell.Interior.ColorIndex = lcol Then
            If rcell.Interior.Color = RGB(0, 127, 0) Or rcell.Interior.Color = RGB(254, 0, 0) Or rcell.Interior.Color = RGB(254, 254, 0) Then
                vresult = 1 + vresult
            End If
        Next rcell
    End If
 End If
 colorfunction = vresult
 

End Function



