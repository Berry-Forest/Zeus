VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} a_Thinking1 
   Caption         =   "Theorizing"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   OleObjectBlob   =   "A_Thinking1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "A_Thinking1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

Me.BwsrEinstein.Navigate ("C:\Users\" & Usr & "\Desktop\Zeus\1-Tools\Components\Gifs\Einstein\Einstein.gif")
Me.BwsrEinstein.Width = 475
Me.Width = 475

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height / 2
Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2

End Sub
