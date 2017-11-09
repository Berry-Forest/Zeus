VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} a_Chuck 
   Caption         =   "UserForm1"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5730
   OleObjectBlob   =   "a_Chuck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "a_Chuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height / 2
Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2

End Sub
