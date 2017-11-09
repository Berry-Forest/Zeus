VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BulkEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents SetupEvents As MSForms.CommandButton
Attribute SetupEvents.VB_VarHelpID = -1
Public WithEvents NtwkChkBoxEvents As MSForms.CheckBox
Attribute NtwkChkBoxEvents.VB_VarHelpID = -1

Public WithEvents PSCenter As MSForms.ComboBox
Attribute PSCenter.VB_VarHelpID = -1
Public WithEvents AddContract As MSForms.CommandButton
Attribute AddContract.VB_VarHelpID = -1
Public WithEvents AddContractBox As MSForms.TextBox
Attribute AddContractBox.VB_VarHelpID = -1
Public WithEvents RemoveReport As MSForms.CommandButton
Attribute RemoveReport.VB_VarHelpID = -1
Public WithEvents RemoveContract As MSForms.CommandButton
Attribute RemoveContract.VB_VarHelpID = -1
Public WithEvents RemoveAsscContract As MSForms.ComboBox
Attribute RemoveAsscContract.VB_VarHelpID = -1

Dim conn As New ADODB.Connection
Sub RemoveReport_Click()

FrameNmbr = Replace(RemoveReport.Name, "ReportRemove", "")

BulkReportSetup.Controls.Remove "ReportFrame" & FrameNmbr

'Adjust Frames
'-----------------
For Each c In BulkReportSetup.Controls
    If TypeName(c) = "Frame" And Replace(c.Name, "ReportFrame", "") > FrameNmbr Then
        c.Top = c.Top - c.Height - 5
        ReportNmbr = Val(Replace(c.Caption, "Report", ""))
        c.Caption = "Report" & ReportNmbr - 1
    End If
Next

'Adjust Arrays
'-----------------
For i = FrameNmbr + 1 To ReportNmbr
    Set PSCArray(i - 1) = PSCArray(i)
    Set AddConArray(i - 1) = AddConArray(i)
    Set RmvConArray(i - 1) = RmvConArray(i)
    Set RmvAConArray(i - 1) = RmvConArray(i)
    Set AddConBxArray(i - 1) = AddConBxArray(i)
    Set RmvReport(i - 1) = RmvReport(i)
Next
ReportNmbr = ReportNmbr - 1

ReDim Preserve PSCArray(1 To ReportNmbr)
ReDim Preserve AddConArray(1 To ReportNmbr)
ReDim Preserve RmvConArray(1 To ReportNmbr)
ReDim Preserve RmvAConArray(1 To ReportNmbr)
ReDim Preserve AddConBxArray(1 To ReportNmbr)
ReDim Preserve RmvReport(1 To ReportNmbr)


End Sub
Private Sub RemoveAsscContract_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    FrameNmbr = Replace(RemoveAsscContract.Name, "AsscContracts", "")
    Set ConCtrl = BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("AsscContracts" & FrameNmbr)
    
    For i = ConCtrl.ListCount - 1 To 0 Step -1
        If ConCtrl.List(i) = ConCtrl.Value Then ConCtrl.RemoveItem i
    Next
    
    If ConCtrl.ListCount = 0 Then
        ConCtrl.Value = ""
    Else
        ConCtrl.Value = ConCtrl.List(0)
    End If

End Sub
Private Sub RemoveContract_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    FrameNmbr = Replace(RemoveContract.Name, "ContractRemove", "")
    Set ConCtrl = BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("AsscContracts" & FrameNmbr)
    
    For i = 0 To ConCtrl.ListCount - 1
        ConCtrl.RemoveItem 0
    Next
    ConCtrl.Value = ""

End Sub
Private Sub RemoveContract_Click()
    
    FrameNmbr = Replace(RemoveContract.Name, "ContractRemove", "")
    Set ConCtrl = BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("AsscContracts" & FrameNmbr)
    
    For i = ConCtrl.ListCount - 1 To 0 Step -1
        If ConCtrl.List(i) = ConCtrl.Value Then ConCtrl.RemoveItem i
    Next
    
    If ConCtrl.ListCount = 0 Then
        ConCtrl.Value = ""
    Else
        ConCtrl.Value = ConCtrl.List(0)
    End If

End Sub
Private Sub AddContract_Click()
    
    FrameNmbr = Replace(AddContract.Name, "ContractAdd", "")
    Set ConCtrl = BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("AsscContracts" & FrameNmbr)
    
    If BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("FromSelectionChk" & FrameNmbr).Value = True Then
        For Each c In Selection
            If Not Trim(c.Value) = "" Then ConCtrl.AddItem Trim(c.Value)
        Next
    Else
        ConCtrl.AddItem Trim(BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("ContractAddBox" & FrameNmbr).Value)
        ConCtrl.Value = Trim(BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("ContractAddBox" & FrameNmbr).Value)
        BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("ContractAddBox" & FrameNmbr).Value = ""
    End If
    ConCtrl.Value = ConCtrl.List(0)


End Sub
Sub AddContractBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)
        
    FrameNmbr = Replace(AddContractBox.Name, "ContractAddBox", "")
    Set ConCtrl = BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("AsscContracts" & FrameNmbr)
    ConCtrl.AddItem Trim(AddContractBox.Value)
    ConCtrl.Value = Trim(AddContractBox.Value)
    AddContractBox.Value = ""

End Sub
Sub PSCenter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


Dim recset As New ADODB.Recordset

    If Not KeyCode = 13 Then Exit Sub  '(trigger event on Enter, Tab=9)

    FrameNmbr = Replace(PSCenter.Name, "AsscPSC", "")
    Set ConCtrl = BulkReportSetup.Controls("ReportFrame" & FrameNmbr).Controls("AsscContracts" & FrameNmbr)
    
    For i = 0 To ConCtrl.ListCount - 1
        ConCtrl.RemoveItem 0
    Next
    
    sqlstr = "" & _
                "SELECT " & _
                    "con.CONTRACT_NUMBER, " & _
                    "con.CONTRACT_EFF_DATE, " & _
                    "con.VENDOR_KEY " & _
                "FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con " & _
                "INNER JOIN OCSDW_GPO_CONTRACT gpo " & _
                    "ON con.CONTRACT_ID = gpo.CONTRACT_ID " & _
                    "AND gpo.COMPANY_CODE = '" & FUN_CmpyCd(BulkReportSetup.Controls("NtwkLbl").Caption) & "' " & _
                "WHERE " & _
                    "con.ATTRIBUTE_VALUE_NAME = '" & PSCenter.Value & "' " & _
                    "AND con.STATUS_KEY IN ('ACTIVE','SIGNED','PENDING') " & _
                    "AND con.EXPORT_TYPE_KEY = 'M' ORDER BY VENDOR_KEY, CONTRACT_EFF_DATE DESC"
       
    If conn = "" Then conn.Open "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
    recset.Open sqlstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic

    If recset.RecordCount = 1 Then
        ConCtrl.AddItem Trim(recset.Fields(0))
    Else
        Set ConSort = New ADODB.Recordset
        ConSort.Fields.Append "Cons", adVarChar, 20
        ConSort.Open
        ConSort.AddNew
        ConSort.Fields(0) = recset.Fields(0)
        prevVndr = recset.Fields(2)
        For i = 1 To recset.RecordCount - 1
            recset.MoveNext
            If Not recset.Fields(2) = prevVndr Then
                ConSort.AddNew
                ConSort.Fields(0) = recset.Fields(0)
                prevVndr = recset.Fields(2)
            End If
        Next
        ConSort.Sort = "Cons"
        For i = 1 To ConSort.RecordCount
            ConCtrl.AddItem Trim(ConSort.Fields(0))
            ConSort.MoveNext
        Next
        Set ConSort = Nothing
    End If
    ConCtrl.Value = ConCtrl.List(0)

    recset.Close
    
    'asscPSC.SetFocus


End Sub
Private Sub NtwkChkBoxEvents_Click()

ntwk = Replace(NtwkChkBoxEvents.Name, "Chk", "")
If NtwkChkBoxEvents.Value = True Then
    NetworkSelection.Controls(ntwk & "Setup").Enabled = True
Else
    NetworkSelection.Controls(ntwk & "Setup").Enabled = False
End If


End Sub
Sub SetupEvents_Click()

ntwk = Replace(SetupEvents.Name, "Setup", "")
BulkReportSetup.Show (False)
BulkReportSetup.NtwkLbl.Caption = ntwk


End Sub

