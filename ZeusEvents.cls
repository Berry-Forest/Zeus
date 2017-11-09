VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZeusEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public WithEvents XLApp As Excel.Application
'Dim mXLHwnd As Long    'Excel's window handle
'Dim mhwndForm As Long  'The userform's window handle
'Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Const GWL_HWNDPARENT As Long = -8
'If Val(Application.Version) >= 15 Then        'Only makes sense on Excel 2013 and up
'    Set XLApp = Application
'    mhwndForm = FindWindowA("Zeus", Caption)
'End If

Public WithEvents CUremoveEvents As MSForms.CommandButton
Attribute CUremoveEvents.VB_VarHelpID = -1
Private Sub CUremoves_click()

CUnmbr = CUnmbr - 1

CUpos = Replace(CUremoves.Name, "CUremove", "")
ZeusForm.Controls.Remove "CUorgnl" & CUpos
ZeusForm.Controls.Remove "CUchg" & CUpos
ZeusForm.Controls.Remove "CUmfg" & CUpos
ZeusForm.Controls.Remove "CUto" & CUpos

If Not CUnmbr = 0 Then
    For Each Ctrl In ConvertUOMsFrame.Controls
        CtrlPos = Right(Ctrl.Name, 1)
        If IsNumeric(CtrlPos) And CtrlPos > CUpos Then Ctrl.Top = Ctrl.Top - CUremoves.Height
    Next
End If

ConvertUOMsFrame.Height = ConvertUOMsFrame.Height - CUremoves.Height
ZeusForm.Controls.Remove "CUremove" & CUpos

End Sub

