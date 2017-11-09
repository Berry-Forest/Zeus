VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AltPRSform 
   Caption         =   "Alt PRS"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3090
   OleObjectBlob   =   "AltPRSform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AltPRSform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    Set LblModel = AltPRSform.Controls("AltPRSlbl1")
    Set TxtModel = AltPRSform.Controls("AltPRScon1")
    Set TxtAdd = TxtModel
    TxtModel.Text = ZeusForm.asscContracts.List(0)
    
    
    For i = 2 To suppNMBR
        
        Set LblAdd = AltPRSform.Controls.Add("Forms.Label.1", "AltPRSlbl" & i, True)
        With LblAdd
            .Caption = "Supplier " & i & ":"
            .Width = LblModel.Width
            .Height = LblModel.Height
            .Left = LblModel.Left
            .Top = LblModel.Top + LblModel.Height * (i - 1)
            .BackStyle = LblModel.BackStyle
            .Font.Bold = LblModel.Font.Bold
        End With
        
        Set TxtAdd = AltPRSform.Controls.Add("Forms.textbox.1", "AltPRScon" & i, True)
        With TxtAdd
            .Text = ZeusForm.asscContracts.List(i - 1)
            .Width = TxtModel.Width
            .Height = TxtModel.Height
            .Left = TxtModel.Left
            .Top = TxtModel.Top + TxtModel.Height * (i - 1)
            .BackStyle = TxtModel.BackStyle
        End With
        
    Next
    
    
    AltPRSok.Top = TxtAdd.Top + TxtAdd.Height + 10
    Me.Height = AltPRSok.Top + AltPRSok.Height + 25
    

End Sub
Sub AltPRSok_Click()

    Me.Hide

End Sub
