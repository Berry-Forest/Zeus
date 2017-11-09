VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BulkReportSetup 
   Caption         =   "Setup"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   OleObjectBlob   =   "BulkReportSetup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BulkReportSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub AddPSC_Click()

    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then ReportNmbr = ReportNmbr + 1
    Next
    ReportNmbr = ReportNmbr + 1
    
    'add section frame
    '------------------------
    Set Ctrl = Me.Controls("ReportFrame1")
    Set FrameAdd = Me.Controls.Add("forms.frame.1", "ReportFrame" & ReportNmbr, True)
    With FrameAdd
        .Caption = "Report" & ReportNmbr
        .Font.Bold = Ctrl.Font.Bold
        .Font.Underline = Ctrl.Font.Underline
        '.BackColor = CtrlModel.BackColor
        .BorderStyle = Ctrl.BorderStyle
        .BorderColor = Ctrl.BorderColor
        '.BackStyle = MemberFrame1.BackStyle
        .Height = Ctrl.Height
        .Width = Ctrl.Width
        .Left = Ctrl.Left
        .Top = Ctrl.Top + (Ctrl.Height + 5) * (ReportNmbr - 1)
    End With
    
    On Error Resume Next
    For Each Ctrl In ReportFrame1.Controls
        Set CtrlAdd = FrameAdd.Controls.Add("Forms." & TypeName(Ctrl) & ".1", Left(Ctrl.Name, Len(Ctrl.Name) - 1) & ReportNmbr, True)
        With CtrlAdd
            .Caption = Ctrl.Caption
            .Font.Bold = Ctrl.Font.Bold
            .Font.Underline = Ctrl.Font.Underline
            .BorderStyle = Ctrl.BorderStyle
            .BorderColor = Ctrl.BorderColor
            .Width = Ctrl.Width
            .Height = Ctrl.Height
            .Left = Ctrl.Left
            .Top = Ctrl.Top
        End With
    Next
    
    ReDim Preserve PSCArray(1 To ReportNmbr)
    Set PSCArray(ReportNmbr).PSCenter = Me.Controls("asscpsc" & ReportNmbr)
    
    ReDim Preserve AddConArray(1 To ReportNmbr)
    Set AddConArray(ReportNmbr).AddContract = Me.Controls("ContractAdd" & ReportNmbr)
    
    ReDim Preserve AddConBxArray(1 To ReportNmbr)
    Set AddConBxArray(ReportNmbr).AddContractBox = Me.Controls("ContractAddBox" & ReportNmbr)
    
    ReDim Preserve RmvConArray(1 To ReportNmbr)
    Set RmvConArray(ReportNmbr).RemoveContract = Me.Controls("ContractRemove" & ReportNmbr)
    
    ReDim Preserve RmvAConArray(1 To ReportNmbr)
    Set RmvAConArray(ReportNmbr).RemoveAsscContract = Me.Controls("AsscContracts" & ReportNmbr)
    
    ReDim Preserve RmvReport(1 To ReportNmbr)
    Set RmvReport(ReportNmbr).RemoveReport = Me.Controls("ReportRemove" & ReportNmbr)

    For i = 0 To Me.Controls("asscPSC1").ListCount - 1
        Me.Controls("asscPSC" & ReportNmbr).AddItem Me.Controls("ReportFrame1").Controls("asscPSC1").List(i)
    Next
            
    Me.Height = FrameAdd.Top + FrameAdd.Height + 20


End Sub

Sub UserForm_Initialize()



    ReDim PSCArray(1 To 1)
    Set PSCArray(1).PSCenter = Me.Controls("asscpsc1")
    
    ReDim AddConArray(1 To 1)
    Set AddConArray(1).AddContract = Me.Controls("ContractAdd1")
    
    ReDim AddConBxArray(1 To 1)
    Set AddConBxArray(1).AddContractBox = Me.Controls("ContractAddBox1")
    
    ReDim RmvConArray(1 To 1)
    Set RmvConArray(1).RemoveContract = Me.Controls("ContractRemove1")
    
    ReDim RmvAConArray(1 To 1)
    Set RmvAConArray(1).RemoveAsscContract = Me.Controls("AsscContracts1")
    
    ReDim RmvReport(1 To 1)
    Set RmvReport(1).RemoveReport = Me.Controls("ReportRemove1")
    
    
    'populate PSCs
    '-------------------
    For i = 0 To ZeusForm.asscPSC.ListCount - 1
        Me.Controls("asscPSC1").AddItem ZeusForm.asscPSC.List(i)
    Next
    
    Me.Height = Me.Controls("ReportFrame1").Top + Me.Controls("ReportFrame1").Height + 21

    Me.StartUpPosition = 0
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2



End Sub
Sub userform_terminate()

ReDim PSCArray(0 To 0)
ReDim AddConArray(0 To 0)
ReDim RmvConArray(0 To 0)
ReDim RmvAConArray(0 To 0)
ReDim AddConBxArray(0 To 0)
ReDim RmvReport(0 To 0)


End Sub
Sub okBttn_click()

'add info to jagged array

'Call gatherBulk

    Dim TempArray() As String

    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then ReportNmbr = ReportNmbr + 1
    Next

    ReDim TempArray(1 To ReportNmbr, 0 To 10)
    For i = 1 To ReportNmbr
        TempArray(i, 0) = Me.Controls("ReportFrame" & i).Controls("asscPSC" & i).Value
        For con = 0 To Me.Controls("ReportFrame" & i).Controls("AsscContracts" & i).ListCount - 1
            TempArray(i, con + 1) = Me.Controls("ReportFrame" & i).Controls("AsscContracts" & i).List(con)
        Next
    Next

    For i = 1 To UBound(NtwkNmArray)
        If NtwkNmArray(i) = NtwkLbl.Caption Then ReportArray(i) = TempArray
    Next

    Unload Me

End Sub


