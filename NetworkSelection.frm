VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NetworkSelection 
   Caption         =   "Network Selection"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10275
   OleObjectBlob   =   "NetworkSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NetworkSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SetupArray() As New BulkEventHandler
Dim ChkBoxArray() As New BulkEventHandler
Private Sub TeamSelect_Change()

    TeamStr = FUN_ConvGroups(AdminconfigStr, TeamSelect.Value & " Networks")
    For i = 1 To UBound(NtwkNmArray)
        If InStr(TeamStr, NtwkNmArray(i)) > 0 Then
            Me.Controls(NtwkNmArray(i) & "Chk").Value = True
            Me.Controls(NtwkNmArray(i) & "Setup").Enabled = True
        Else
            Me.Controls(NtwkNmArray(i) & "Chk").Value = False
            Me.Controls(NtwkNmArray(i) & "Setup").Enabled = False
        End If
    Next
    

End Sub
Private Sub okBttn_click()

    Me.Hide
    
    'Dim ConArray() As String
    'Dim NewApp As New Excel.Application
    
    For ntwk = 1 To UBound(ReportArray)
        On Error GoTo ERR_NoNtwk
        If Not Trim(ReportArray(ntwk)(1, 0)) = "" Then
            ReportNtwk = NtwkNmArray(ntwk)
            For Rprt = 1 To UBound(ReportArray(ntwk))
                Constr = ""
                For i = 0 To 10 'UBound(ReportArray(ntwk)(Rprt)) - 1
                    If i = 0 Then
                        ReportPSC = ReportArray(ntwk)(Rprt, i)
                    Else
                        If Not Trim(ReportArray(ntwk)(Rprt, i)) = "" Then
                            'ReDim Preserve ConArray(1 To i)
                            'ConArray(i) = ReportArray(ntwk)(Rprt, i)
                            Constr = Constr & """ """ & ReportArray(ntwk)(Rprt, i)
                        End If
                    End If
                Next
                
                Constr = Mid(Constr, 4, Len(Constr)) & """"
                ScriptDir = Chr(34) & EnvironPath & "\Admin\Scripts\RunZeusReport.vbs" & Chr(34)
                Shell "wscript " & ScriptDir & " """ & ReportNtwk & """ """ & ReportPSC & """ """ & Constr
                
                'NewApp.Visible = True
                'Set NewZeus = NewApp.Workbooks.Open(FUN_ConvTags(AdminconfigStr, "Local App Path") & "\" & FUN_ConvTags(AdminconfigStr, "App"))
                'NewApp.Run "BulkRun", NetVar, PSCVar, ConArray
                
            Next
        End If
nxtNtwk:
    Next
    
    
Exit Sub
':::::::::::::::::::::::
ERR_NoNtwk:
Resume nxtNtwk



End Sub
Private Sub SelectAllChk_Click()


    If SelectAllChk = True Then
        BooVal = True
    Else
        BooVal = False
    End If
    
    For i = 1 To UBound(NtwkNmArray)
        Me.Controls(NtwkNmArray(i) & "Chk").Value = BooVal
        Me.Controls(NtwkNmArray(i) & "Setup").Enabled = BooVal
    Next


End Sub
Private Sub UserForm_Initialize()

On Error Resume Next

    FontSz = 20
    FontNm = "Impact" '"Gill Sans Ultra Bold" '"Cooper Black" ' ' '
    FontBld = False
    FontClr = &HFFFF& '&H8080FF  '&H80C0FF '&HFF00& '&H80FFFF  '  ' ' ''  ' '   ' ' ' ' '         '         '    '
    MaxLblWidth = FontSz * 2
    MaxLblHeight = FontSz + 5
    MaxRdlWidth = FontSz * 3
    MaxRdlHeight = FontSz * 1.5
    MaxChkHeight = FontSz
    MaxBttnWidth = FontSz
    MaxBttnHeight = FontSz '* 0.75
    
    Dim DataSources(1 To 3) As String
    DataSources(1) = "RDM"
    DataSources(2) = "NR"
    DataSources(3) = "Extract"

    SNAstr = FUN_ConvGroups(AdminconfigStr, "SNA Networks")
    CustomStr = FUN_ConvGroups(AdminconfigStr, "CSA Networks")

    For ntwk = 1 To UBound(NtwkNmArray)
    
        'Add ntwk checkbox
        '----------------------
        Set ChkAdd = NetworkSelection.Controls.Add("Forms.checkbox.1", NtwkNmArray(ntwk) & "Chk", True)
        With ChkAdd
            '.Value = True
            .Width = 12
            .Height = 12
            .Left = 12
            .Top = 20 + (MaxChkHeight + 5) * ntwk '20 + ChkAdd.Height * ntwk '+ CurrHeight
            .BackStyle = 0
        End With
        ReDim Preserve ChkBoxArray(1 To ntwk)
        Set ChkBoxArray(ntwk).NtwkChkBoxEvents = ChkAdd
        
        'add ntwk name label
        '----------------------
        Set LblAdd = NetworkSelection.Controls.Add("Forms.Label.1", NtwkNmArray(ntwk) & "Lbl", True)
        With LblAdd
            .Caption = NtwkNmArray(ntwk)
            .Width = MaxLblWidth * 2
            .Height = MaxLblHeight
            '.AutoSize
            .Left = ChkAdd.Left + ChkAdd.Width
            .Top = ChkAdd.Top - ChkAdd.Height / 2 '- LblAdd.Height / 4 '+ ChkAdd.Height / 2 - LblAdd.Height / 4
            .BackStyle = 0
            .ForeColor = FontClr
            .Font.Bold = FontBld
            .Font.Name = FontNm
            .Font.Size = FontSz
            '.BorderStyle = 1
            '.BorderColor = &H80FF&
        End With
    
        'Add Source radials
        '----------------------
        For src = 1 To UBound(DataSources)
            Set RdlAdd = NetworkSelection.Controls.Add("Forms.OptionButton.1", NtwkNmArray(ntwk) & DataSources(src) & "Rdl", True)
            With RdlAdd
                If src = 1 And InStr(SNAstr, NtwkNmArray(ntwk)) > 0 Then
                    .Value = True
                ElseIf src = 3 And InStr(CustomStr, NtwkNmArray(ntwk)) > 0 Then
                    .Value = True
                End If
                .Caption = DataSources(src)
                .GroupName = NtwkNmArray(ntwk) & "source"
                .Width = MaxRdlWidth + MaxRdlWidth * 0.5
                .Height = MaxRdlHeight '16
                .Left = LblAdd.Left + LblAdd.Width + RdlAdd.Width * (src - 1)
                .Top = ChkAdd.Top - ChkAdd.Height / 2
                .BackStyle = 0
                .ForeColor = FontClr
                .Font.Name = FontNm
                .Font.Size = FontSz
                .Font.Bold = FontBld
                '.TextAlign = 2
                '.BorderStyle = 1
                '.BorderColor = &H80000012
            End With
        Next
        
        'add ntwk Setup bttn
        '----------------------
        Set SetupAdd = NetworkSelection.Controls.Add("Forms.commandbutton.1", NtwkNmArray(ntwk) & "Setup", True)
        With SetupAdd
            .Caption = "Setup"
            .Width = MaxBttnWidth * 2
            .Height = MaxBttnHeight
            '.AutoSize
            .Left = RdlAdd.Left + RdlAdd.Width
            .Top = ChkAdd.Top '- ChkAdd.Height / 2 '- LblAdd.Height / 4 '+ ChkAdd.Height / 2 - LblAdd.Height / 4
            '.ForeColor = FontClr
            .Font.Bold = FontBld
            '.Font.Name = FontNm
            .Font.Size = 8
            .Enabled = False
        End With
        ReDim Preserve SetupArray(1 To ntwk)
        Set SetupArray(ntwk).SetupEvents = SetupAdd
    
    Next
    
    Me.Width = SetupAdd.Left + SetupAdd.Width + 10 '20 + chkAdd.Width + LblAdd.Width + UBound(DataSources) *
    Me.Height = RdlAdd.Top + RdlAdd.Height + Okbttn.Height + 33
    Okbttn.Left = Me.Width / 2 - Okbttn.Width / 2
    Okbttn.Top = RdlAdd.Top + RdlAdd.Height + 7
    SelectAllLbl.Font.Name = FontNm
    SelectAllLbl.ForeColor = FontClr
    SelectAllChk.Value = False

    TeamSelect.AddItem "SNA"
    TeamSelect.AddItem "CSA"
    
    ReDim ReportArray(1 To UBound(NtwkNmArray))

    Me.StartUpPosition = 0
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2

End Sub
