VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateXrefForm 
   Caption         =   "Create Xref"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   OleObjectBlob   =   "CreateXrefForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateXrefForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormCatnum_Enter()

On Error Resume Next
Set txtRng = Application.InputBox("Please select column with associated catalog numbers", Type:=8)
Sheets(txtRng.Parent.Name).Select
If Err <> 0 Then
    Me.FormCatnum.Text = ""
Else
    Me.FormCatnum.Text = "'" & txtRng.Parent.Name & "'!" & Range(Cells(1, txtRng.Column), Cells(FUN_lastrow(txtRng.Column), txtRng.Column)).Address
End If
Me.FormName.SetFocus


End Sub
Private Sub FormDesc_Enter()

On Error Resume Next
Set txtRng = Application.InputBox("Please select column with associated descriptions", Type:=8)
Sheets(txtRng.Parent.Name).Select
If Err <> 0 Then
    Me.FormDesc.Text = ""
Else
    Me.FormDesc.Text = "'" & txtRng.Parent.Name & "'!" & Range(Cells(1, txtRng.Column), Cells(FUN_lastrow(txtRng.Column), txtRng.Column)).Address
End If
Me.FormName.SetFocus


End Sub
Private Sub XrefOk_Click()

If Me.FormCatnum.Text = "" Then
    MsgBox "Please enter a value for This Supplier Catalog Numbers"
ElseIf Trim(FormName.Text) = "" Then
    MsgBox "Please enter a value for Supplier Name"
Else
    Call ExtractSupplierXref
    Me.FormCatnum.Text = ""
    Me.FormDesc.Text = ""
    Me.FormName.Text = ""
    Me.XrefInstructions.Caption = "Please enter next supplier name and select columns"
    Me.XrefInstructions.Left = 6
End If

End Sub
Private Sub XrefDone_Click()

Call EndCreateXref
'Me.FormCatnum.Text = ""
'Me.FormDesc.Text = ""
'Me.FormName.Text = ""
Unload CreateXrefForm

End Sub
Private Sub userform_terminate()

Set CreateXrefWB = Nothing

End Sub
