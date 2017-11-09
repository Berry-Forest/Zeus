VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QCform 
   Caption         =   "QC Checklist"
   ClientHeight    =   18585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13020
   OleObjectBlob   =   "QCform.frx":0000
End
Attribute VB_Name = "QCform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IMmbrs As Range
Public lnkCTRL As control
Public txtCTRL As control

Dim BoxArray() As New QCEventHandler
Dim NoteArray() As New QCEventHandler
Dim CheckDesc(1 To 49) As String
Private Sub userform_terminate()

ReviewFlg = 0

End Sub
Private Sub UserForm_Activate()

AddToForm MIN_BOX

End Sub
Private Sub UserForm_Initialize()


    Dim HdrTitles(1 To 12) As String
    HdrTitles(1) = "Notes_2"
    HdrTitles(2) = "Spend Totals Match On:_6"
    HdrTitles(3) = "Index_6"
    HdrTitles(4) = "Initiative Spend Overview_3"
    HdrTitles(5) = "Graphs_4"
    HdrTitles(6) = "Vizient Contracts - Conv_1"
    HdrTitles(7) = "Line Item Data_13"
    HdrTitles(8) = "Pricing_5"
    HdrTitles(9) = "Cross References_2"
    HdrTitles(10) = "Admin Fees_1"
    HdrTitles(11) = "Best Market Price_1"
    HdrTitles(12) = "Overall_5"
    
    
    
    'Check Descriptions
    '------------------------------------------------
    CheckDesc(1) = "DAT name included"
    CheckDesc(2) = "Variance ranges included"
    CheckDesc(3) = "1) Market Share"
    CheckDesc(4) = "2) Benchmark"
    CheckDesc(5) = "3) PRS"
    CheckDesc(6) = "4) Non Conversion"
    CheckDesc(7) = "5) Conversion"
    CheckDesc(8) = "6) Line Item Data"
    CheckDesc(9) = "Report Create Date"
    CheckDesc(10) = "Member date ranges accurate"
    CheckDesc(11) = "Required members present and standardized"
    CheckDesc(12) = "UNSPSC data present"
    CheckDesc(13) = "Tier data and descriptions are accurate"
    CheckDesc(14) = "Items on contract accurate"
    CheckDesc(15) = "Member spend totals are reasonable"
    CheckDesc(16) = "PRS data is present and reasonable"
    CheckDesc(17) = """All Others"" total <= to 5%"
    CheckDesc(18) = "Market share graph matches data"
    CheckDesc(19) = "Market share graph is readable"
    CheckDesc(20) = "Benchmarking graph matches data"
    CheckDesc(21) = "Benchmarking graph is readable"
    CheckDesc(22) = "Proposed contract totals sum correctly"
    CheckDesc(23) = "Contracted Mftr names match MPP supplier name"
    CheckDesc(24) = "Non contracted Mftr names standardized"
    CheckDesc(25) = "Only one Mftr name per catalog number"
    CheckDesc(26) = "No blank, unknown, or distributor Mftr names"
    CheckDesc(27) = "Catalog numbers standardized"
    CheckDesc(28) = "No blank or unknown catalog numbers"
    CheckDesc(29) = "Price leveling variances have been resolved"
    CheckDesc(30) = "Tenth percentile variances resolved"
    CheckDesc(31) = "Supplier variances have been resolved"
    CheckDesc(32) = "Supplier Benchmark variances resolved"
    CheckDesc(33) = "UOM pkg descriptions and UOM quantities match"
    CheckDesc(34) = "Out of scope items removed"
    CheckDesc(35) = "Convert to Novaplus codes where applicable"
    CheckDesc(36) = "Duplicate catalog numbers reconciled"
    CheckDesc(37) = "Tier Used formulas correct"
    CheckDesc(38) = "$0 price items removed"
    CheckDesc(39) = "Unqualified tiers removed"
    CheckDesc(40) = "Verified most current pricing being used"
    CheckDesc(41) = "Data sorted correctly"
    CheckDesc(42) = "Xref codes cleansed"
    CheckDesc(43) = "Correct Admin Fees are used"
    CheckDesc(44) = "Data correctly sorted"
    CheckDesc(45) = "Report Fully Caluclated"
    CheckDesc(46) = "All fonts Arial 8pt and zooms at 100%"
    CheckDesc(47) = "Original extract, ASF, and Xref posted"
    CheckDesc(48) = "Member data intact vs original extract"
    CheckDesc(49) = "No hardcoded formulas in dedicated formula cells"
    
    'populate Frames/Headers, desc labels, and status boxes,
    '=====================================================================================================
    LblWdth = 180
    BoxWdth = 12
    FrameWdth = LblWdth + BoxWdth + 4
    FormHght = 30
    
    For i = 1 To UBound(HdrTitles)
        HdrStr = HdrTitles(i)
        
        'Add Frame
        '--------------------------------
        Set FrameAdd = Me.Controls.Add("forms.frame.1", "Frame" & i, True)
        With FrameAdd
            .Top = FormHght + 4
            .Left = 2
            '.Height = CtrlModel.Height
            .Width = FrameWdth
            .Caption = Left(HdrStr, InStr(HdrStr, "_") - 1)
            .Font.Bold = True
            .Font.Italic = False
            .Font.Size = 8
            '.BackStyle = 0
            .BackColor = &HFFFFFF
            .BorderStyle = 1
        End With
        
        'add checkbox and label
        '--------------------------------
        FrameHght = 6
        For j = 1 To Mid(HdrStr, InStr(HdrStr, "_") + 1, Len(HdrStr))
            CurrItm = CurrItm + 1
            Set LabelAdd = FrameAdd.Controls.Add("Forms.label.1", "Label" & CurrItm, True)
            With LabelAdd
                .Top = FrameHght
                .Left = 3
                .Height = 12
                .Width = LblWdth
                .Caption = CheckDesc(CurrItm)
                .Font.Bold = False
                .Font.Italic = False
                .Font.Size = 8
                .BackStyle = 0
            End With
            
            Set BoxAdd = FrameAdd.Controls.Add("Forms.commandbutton.1", "StatusBox" & CurrItm, True)
            With BoxAdd
                .Top = FrameHght
                .Left = LabelAdd.Left + LabelAdd.Width
                .Height = 12
                .Width = BoxWdth
                .Caption = vbNullString
                .BackColor = 65535
            End With
            ReDim Preserve BoxArray(1 To CurrItm)
            Set BoxArray(CurrItm).BoxEvents = BoxAdd
            
            Set NoteAdd = FrameAdd.Controls.Add("Forms.label.1", "Note" & CurrItm, True)
            With NoteAdd
                .Top = BoxAdd.Top
                .Left = BoxAdd.Left - 4.5
                .Height = 9
                .Width = 4.5
                .Caption = "!"
                .Font.Bold = True
                .Font.Italic = True
                .Font.Size = 8
                .BackStyle = 0
                .Visible = False
            End With
            ReDim Preserve NoteArray(1 To CurrItm)
            Set NoteArray(CurrItm).NoteEvents = NoteAdd
            
            FrameHght = FrameHght + LabelAdd.Height + 1.5
        Next
        FrameAdd.Height = FrameHght + 7
        FormHght = FrameAdd.Top + FrameAdd.Height
    Next
    
    Me.Width = FrameWdth + 25
    NoteFrame.Height = FormHght + 20
    NoteLabel.Height = FormHght + 20
    NoteFrame.Width = FrameWdth + 3
    NoteLabel.Width = FrameWdth - 2
    NoteFrame.Top = 2
    NoteFrame.Left = 2
    NoteLabel.Top = 18
    NoteLabel.Left = 0
    
    
    'set scrollbar properties
    '------------------------------------
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = Me.Height - 30
    Me.Height = Application.Height - 30
    
    
    AddToForm MIN_BOX
    
    
End Sub
Sub NoteCloseBttn_Click()

    QCform.Controls("NoteFrame").Visible = False

End Sub
Private Sub PostToI_Click()

    Call PostToIMETH

End Sub

Private Sub PubQC_Click()



    'Check sheets and see if QC tab needs to be created
    '============================================================================================================
    On Error GoTo ERR_NoQC
    Sheets("QC").Visible = True
    Sheets("QC").Select
    On Error GoTo 0
    ActiveSheet.Unprotect Password = "existentialism"
    
    'publish status'
    '============================================================================================================
    On Error Resume Next
    For i = 1 To UBound(CheckDesc)
        Range("C:C").Find(what:=CheckDesc(i)).Offset(0, 1).Interior.Color = Me.Controls("StatusBox" & i).BackColor
    Next

    'Report details
    '============================================================================================================
    net_init = Sheets("Index").Range("C7").Value
    Range("K2").Value = Trim(Left(net_init, InStr(net_init, "-") - 2))
    Range("K3").Value = Mid(net_init, InStr(net_init, "-") + 2, Len(net_init))
    If Not ReviewFlg = 1 Then
        Range("K4").Value = Usr
    Else
        Range("K5").Value = Usr
    End If

    ActiveSheet.Protect Password = "existentialism"


Exit Sub
':::::::::::::::::::::::::::::::::::
ERR_NoQC:
    Call QC_Template
Resume



End Sub
Sub QCsubmitBttn_click()

If Not FUN_Save = vbYes Then Exit Sub
SetupSwitch = FUN_SetupSwitch

Call QCsubmit


End Sub
Sub QCunlockBttn_click()

Call unlockQC(QCunlockBttn)

End Sub
Sub QCHelpBttn_click()

Call QCHelp(QCHelpBttn)

End Sub
