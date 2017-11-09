VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcelEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents App As Application   'associated the "WithEvents" command with this class module so that all events declared in the "this workbook" object are passed to this module.
Attribute App.VB_VarHelpID = -1
Private Sub Class_Initialize()
    Set App = Application
End Sub
'Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
'
''is activated when an different workbook other than the zeus addin is saved. if zeus workbook addin is saved it will initialize another instance of this class module.
''*********************************************************************
'
''Range("A1").Select
''Range("A2").Select
'On Error GoTo defaultexit
'
''refresh common variables
''-----------------------------
'If Success = True Then
'    If Not (IsEmpty(tmWB) Or tmWB Is Nothing) Then
'
'        'search open workbooks
'        '-----------------------------------
'        For Each Wb In Workbooks
'            If InStr(Wb.Name, tmWB.Name) > 0 Then
'
'                'If activewb Is tmWB Then Refresh Variables
'                '-----------------------------------------
'                If ActiveWorkbook.Name = tmWB.Name Then
'                   Call SetCommonVariables  '>>>>>>>>>>>
'                End If
'
'            End If
'        Next
'
'    End If
'End If
'
'defaultexit:
'
'End Sub
'Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
'
'On Error GoTo defaultexit
'If Not IsEmpty(tmWB) And Not tmWB Is Nothing Then
'
'    'search open workbooks
'    '-----------------------------------
'    For Each Wb In Workbooks
'        If InStr(Wb.Name, tmWB.Name) > 0 Then
'
'            'If activewb Is tmWB Then Refresh Variables
'            '-----------------------------------------
'            If ActiveWorkbook.Name = tmWB.Name Then
'                Call destroyVar  '>>>>>>>>>>
'            End If
'
'        End If
'    Next
'
'End If
'
'defaultexit:
'
'End Sub
'Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
'    MsgBox "New Workbook: " & Wb.Name
'End Sub

'Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'
'Dim conn As New ADODB.Connection
'Dim RecSet As New ADODB.Recordset
'
'
'If Sh.Name = "Contract Info" Then
'    If Target.Address = "$N$2" Then
'        If Target.Offset(0, -1).Value = "PSC:" Then
'
'        Application.EnableEvents = False
'
'        Range("B2:K2").ClearContents
'
'        SQLstr = "SELECT CONTRACT_NUMBER FROM OCSDW_CONTRACT_ATTRIBUTE_VALUE_DETAIL con WHERE con.ATTRIBUTE_VALUE_NAME = '" & Range("N2").Value & "' and con.STATUS_KEY = 'ACTIVE' AND con.CONTRACT_NUMBER LIKE '[a-z][a-z][0-9]%[0-9]' ORDER BY CONTRACT_NUMBER"
'
'        ConnStr = "Driver={SQL Server};Server=dwprod.corp.vha.ad;Database=EDB;Trusted_Connection=Yes;"
'        conn.Open ConnStr
'        conn.CommandTimeout = 300
'
'        On Error GoTo errhndlNORECSET
'        RecSet.Open SQLstr, ActiveConnection:=conn, CursorType:=adOpenStatic, LockType:=adLockOptimistic
'        RecSet.MoveFirst
'        On Error GoTo 0
'
'        For j = 1 To RecSet.RecordCount
'            Range("A2").Offset(0, j).Value = RecSet.Fields(0)
'            RecSet.MoveNext
'        Next
'
'        Set RecSet = Nothing
'        Set conn = Nothing
'
'        Application.EnableEvents = True
'
'        End If
'    End If
'End If
'
'
'Exit Sub
'':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'errhndlNORECSET:
'Range("B2").Value = "No contracts found for this PSC"
'Application.EnableEvents = True
'Set conn = Nothing
'Set RecSet = Nothing
'Exit Sub
'
'
'End Sub
