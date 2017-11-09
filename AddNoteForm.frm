VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddNoteForm 
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "AddNoteForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddNoteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddNotePos
Private Sub CopyPrevBttn_Click()

NoteTxt.Value = AddNotePos.Offset(-1, 1).Value
NoteMenuFrame.Visible = False

End Sub
Private Sub NoteCatTitle_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then Call findNote


End Sub
Private Sub NoteTxt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then
    Call noteOK_click
End If

End Sub
Private Sub noteMenu_Click()

If noteMenu = True Then
    NoteMenuFrame.Visible = True
    NoteMenuFrame.Top = noteMenu.Top + noteMenu.Height
    NoteMenuFrame.Left = noteMenu.Left - (NoteMenuFrame.Width - noteMenu.Width)
Else
    NoteMenuFrame.Visible = False
End If


End Sub
Private Sub NoteMenuFrame_Exit(ByVal Cancel As MSForms.ReturnBoolean)

NoteMenuFrame.Visible = False

End Sub

Private Sub UserForm_Activate()

AddToForm MIN_BOX

End Sub

Private Sub UserForm_Click()

NoteTxt.SetFocus
'NoteCatTitle.SetFocus

End Sub

Private Sub UserForm_Deactivate()



End Sub

Private Sub UserForm_Initialize()

NotesOpen = 1
NoteCatTitle.Text = Catnmbr

If tmWB.Sheets("notes").Range("K1").Value = vbNullString Then
    tmWB.Sheets("notes").Range("K1").Value = "Line Item Research"
    tmWB.Sheets("notes").Range("K1").Font.Bold = True
    tmWB.Sheets("notes").Range("K1").Font.Underline = True
End If

'populate user dropdown
'---------------------------
If Not Trim(Sheets("notes").Range("K2").Value) = "" Then
    For Each c In Range(Sheets("notes").Range("K2"), Sheets("notes").Range("K1").End(xlDown))
        For UsrIndex = 0 To UserList.ListCount - 1
            If UserList.List(UsrIndex) = c.Offset(0, -2).Value Then
                GoTo nxtUsr
            End If
        Next
        UserList.AddItem c.Offset(0, -2).Value
nxtUsr:
    Next
    If Not Application.CountIf(Range(Sheets("notes").Range("K2"), Sheets("notes").Range("K1").End(xlDown)).Offset(0, -2), Usr) > 0 Then UserList.AddItem Usr
Else
    UserList.AddItem Usr
End If
Me.UserList = Usr

'populate category dropdown
'---------------------------
With CategoryList
    .AddItem "UOM"
    .AddItem "Clinical"
    .AddItem "QC"
    .AddItem "Data Mining"
End With
If ReviewFlg = 1 Then
    Me.CategoryList = "QC"
ElseIf HansFLG = 1 Then
    Me.CategoryList = "UOM"
ElseIf SherlockFLG = 1 Then
    Me.CategoryList = "Clinical"
ElseIf HermesFLG = 1 Or ExtractFLG = 1 Then
    Me.CategoryList = "Data Mining"
End If

Me.Width = NoteTxt.Left + NoteTxt.Width + 15


End Sub

Sub noteCatTitle_enter()

If ActiveSheet.Name = "Line Item Data" Then NoteCatTitle.Text = ActiveCell.Offset(0, Range("X1").Column - ActiveCell.Column).Value

End Sub
Sub noteOK_click()

If Not NoteTxt.Value = vbNullString Then
    If Not CategoryList.Value = vbNullString Then
        If UserList.Value = Usr Then
            AddNotePos.Value = NoteCatTitle.Text
            AddNotePos.Offset(0, -2).Value = Usr
            AddNotePos.Offset(0, -1).Value = CategoryList.Value
            AddNotePos.Offset(0, 1).Value = NoteTxt.Value
            'AddNotePos.Offset(0, 2).Value = WebRef
            If Not ZeusNotes = 1 Then
                Unload AddNoteForm
            Else
                NoteTxt.Value = ""
            End If
        Else
            MsgBox "You are not " & UserList.Value & ". Please select your name from the dropdown."
        End If
    Else
        MsgBox "Please select a note category from the dropdown."
    End If
End If



End Sub
Private Sub CategoryList_Change()

Call findNote


End Sub

Private Sub userform_terminate()

NotesOpen = 0

End Sub

Private Sub UserList_Change()

Call findNote


End Sub
Sub AddNewBttn_Click()

If Trim(tmWB.Sheets("notes").Range("K2").Value) = "" Then
    Set AddNotePos = tmWB.Sheets("notes").Range("K2")
    NoteTxt.Value = vbNullString
    UserList.Value = Usr
Else
    Set AddNotePos = tmWB.Sheets("notes").Range("K1").End(xlDown).Offset(1, 0)
    NoteTxt.Value = vbNullString
    UserList.Value = Usr
End If

MenuFrame.Visible = False


End Sub
Sub findNote()

If Not tmWB.Sheets("notes").Range("L2").Value = vbNullString Then
    For Each c In Range(tmWB.Sheets("notes").Range("L2"), tmWB.Sheets("notes").Range("L1").End(xlDown)).Offset(0, -1)
        If c.Value = NoteCatTitle.Text Then
            If c.Offset(0, -2) = UserList.Value And c.Offset(0, -1) = CategoryList.Value Then
                Set AddNotePos = c
                NoteTxt.Value = c.Offset(0, 1).Value
                Exit Sub
            End If
        End If
    Next
    Set AddNotePos = tmWB.Sheets("notes").Range("L1").End(xlDown).Offset(1, -1)
    NoteTxt.Value = WebRefNote
Else
    Set AddNotePos = tmWB.Sheets("notes").Range("K2")
    NoteTxt.Value = WebRefNote
    Exit Sub
End If





End Sub

