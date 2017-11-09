Attribute VB_Name = "xxScreenshot"
'Declare Windows API Functions
'---------------------------------
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
  bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
 
'Declare Virtual Key Codes
'---------------------------------
Private Const VK_SNAPSHOT = &H2C
Private Const VK_KEYUP = &H2
Private Const VK_MENU = &H12
Public Const VK_TAB = &H9
Public Const VK_ENTER = &HD
Sub ScreenPrint()

    'Press Alt + TAB Keys
    '---------------------------------
    DoEvents
    keybd_event VK_MENU, 1, 0, 0 'Alt key down
    DoEvents
    keybd_event VK_TAB, 0, 0, 0 'Tab key down
    DoEvents
    keybd_event VK_TAB, 1, VK_KEYUP, 0 'Tab key up
    DoEvents
    keybd_event VK_ENTER, 1, 0, 0 'Tab key down
    DoEvents
    keybd_event VK_ENTER, 1, VK_KEYUP, 0 'Tab key up
    DoEvents
    keybd_event VK_MENU, 1, VK_KEYUP, 0 'Alt key up
    DoEvents
 
    'Press Print Screen key using Windows API
    '---------------------------------
    keybd_event VK_SNAPSHOT, 1, 0, 0 'Print Screen key down
    keybd_event VK_SNAPSHOT, 1, VK_KEYUP, 0 'Print key Up - Screenshot to Clipboard
 
    'Paste Image in Chart and Export it to Image file
    '---------------------------------
    Charts.Add
    ThisWorkbook.Charts(1).AutoScaling = True
    ThisWorkbook.Charts(1).Paste
    ThisWorkbook.Charts(1).Export Filename:="E:\ClipBoardToPic.jpg", FilterName:="jpg"
 
    'Supress warning message and Delete the Chart
    '---------------------------------
    Application.DisplayAlerts = False
    ThisWorkbook.Charts(1).Delete
    Application.DisplayAlerts = True
 

End Sub
Sub PrintScreen_Alt()

'show VB editor
'---------------------------------
VBE.MainWindow.Visible = True
'Application.Goto "Import_PRS"

'Shift-Print Screen
'---------------------------------
Application.SendKeys "(%{1068})"
DoEvents
ActiveWorkbook.Activate
Range("A1").Select
Application.SendKeys "(^v)"

End Sub

Sub splitoff()

''Send in mail
''---------------------------------
'With OutMail
'.To = "youremail@test.com"
'.Subject = "Subject line"
'.display
'Application.SendKeys "(^v)"
'End With

Set clip = CreateObject("clipbrd.clipboard")
SavePicture clip.GetData, "c:\mycliptest.jpg"
'clip.SetData(2, "c:\SavedRange.jpg")
clip.PutInClipboard


End Sub
Sub testre()

Call clipboard.Clear
Call controlRange.execCommand("Copy")
Sheets("Table1").Range("A1").Select
Sheets("Table1").PasteSpecial




End Sub
Sub CopyToClipboard()
    Dim clipboard As MSForms.DataObject
    Dim strSample As String

    Set clipboard = New MSForms.DataObject
    strSample = "This is a sample string"

    
    clipboard.setData "c:\SavedRange.jpg"
    clipboard.SetText strSample
    clipboard.PutInClipboard
    
    clipboard.GetFromClipboard
    str1 = clipboard.GetText
    
End Sub

