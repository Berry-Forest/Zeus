VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} a_Calculating 
   Caption         =   "Calculating...Please Wait"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4410
   OleObjectBlob   =   "a_Calculating.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "a_Calculating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

'Dim queue As Collection

Dim rndmKey As Integer
Dim calcWidth As Integer
Dim calcHeight As Integer

If Not CreateReport = True Then
    Call SetPathVariables
End If

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set queue = New Collection
'queue.aDD objFSO.GetFiles(ZeusMaster.ZeusPATH & "/1-Tools/Components/Gifs")

'Set objFSO = CreateObject("Scripting.FileSystemObject")
setobjFolder (ZeusPATH & "1-Tools/Components/Gifs")

rndmKey = Mid(Time, InStr(Time, " ") - 1, 1)

'Me.WebBrowser1.Navigate (objFolder.Files(rndmKey).Path)
'Debug.Print objFolder.Files(1).Path

rndmSelect = -1
For Each ofile In objFolder.Files
    rndmSelect = rndmSelect + 1
    If rndmSelect = rndmKey Then
        pathStr = ofile.Path
'        Me.Width = LoadPicture(oFile.Path).Width / 25.477
'        Me.Height = LoadPicture(oFile.Path).Height / 25.477
        Me.WebBrowser1.Width = LoadPicture(ofile.Path).Width / 25.477
        Me.WebBrowser1.Height = LoadPicture(ofile.Path).Height / 25.477
        Me.Width = Me.WebBrowser1.Width / 1.3
        Me.Height = Me.WebBrowser1.Height / 1.1
        Exit For
    End If
Next

'objFile.ExtendedProperty ("Dimensions")
Me.WebBrowser1.Navigate (pathStr)

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height / 2
Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2

'Me.WebBrowser1.Navigate ("C:/Users/BFORREST/Desktop/Gentry/1-Tools/Components/Gifs/spinning.gif")
'Me.WebBrowser1.Navigate ("http://images.wikia.com/playstationallstarsbattleroyale/images/0/07/Thumbsup.gif")

End Sub

Private Sub userform_terminate()

'Unload

End Sub
