VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QCEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents BoxEvents As MSForms.CommandButton
Attribute BoxEvents.VB_VarHelpID = -1
Public WithEvents NoteEvents As MSForms.Label
Attribute NoteEvents.VB_VarHelpID = -1

Sub BoxEvents_Click()
    
    CurrCheck = Replace(BoxEvents.Name, "StatusBox", "")
    Call QCChecks(CurrCheck)
    

End Sub
Sub NoteEvents_Click()

If QCform.Controls("NoteFrame").Visible = False Then
    QCform.Controls("NoteFrame").Visible = True
    QCform.Controls("NoteLabel").Caption = Mid(NoteEvents.Caption, InStr(NoteEvents.Caption, " "), Len(NoteEvents.Caption))
    'QCform.Controls("NoteLabel").Caption = "sdfghjk"
    QCform.Controls("NoteFrame").ZOrder (msoBringToFront)
Else
    QCform.Controls("NoteFrame").Visible = False
End If


End Sub


