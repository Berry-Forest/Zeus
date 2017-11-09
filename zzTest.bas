Attribute VB_Name = "zzTest"


Sub arraysubtst()
    
'    Dim conn As New ADODB.Connection
'
'    constr1 = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=exap-scan.corp.vha.com)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=toolsprd_rac)));Uid=bforrest;Pwd=existentialism;"
'    connstr = "Driver={Microsoft ODBC for Oracle};CONNECTSTRING=" & constr1
'    RDMconn.Open connstr

For Each Wb In Application.Workbooks
    Debug.Print Wb.Name
Next


End Sub
Sub arrytst()

configfile = "\\filecluster01\dfs\NovSecure2\SupplyNetworks\Zeus\Prod\Prod1\Get_AdminConfigStr.vbs"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set fileObj = objFSO.GetFile(configfile)
EnvironStr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)

'AconfigStr = Execute EnvironStr



'set usr & environ
'---------------------
Usr = FUN_findUsr
'Usrpath = "C:\Users\" & Usr & "\"
UsrEnviron = UCase(Environ("USERPROFILE")) 'Usrpath
EnvironNmbr = FUN_ConvTags(EnvironStr, "Number of Environments")
For i = 1 To EnvironNmbr
    EnvironUsers = FUN_ConvTags(EnvironStr, "Environ " & i & " Users")
    If InStr(UCase(EnvironUsers), UCase(Usr)) > 0 Then
        AdminConfigPATH = FUN_ConvTags(EnvironStr, "Environ " & i & " Config")
        EnvironPath = FUN_ConvTags(EnvironStr, "Environ " & i & " Path")
        AdminEmail = FUN_ConvTags(EnvironStr, "Admin Email")
        Exit For
    End If
Next

'get adminconfig
'---------------------
Set fileObj = objFSO.GetFile(AdminConfigPATH)
AdminconfigStr = FUN_ConvToStr(fileObj.OpenAsTextStream(1).ReadAll)

Execute str
 


End Sub

Sub testr()

        For i = 1 To suppNMBR
            SuppCol = Sheets("line item data").Range("BG5:BG99999").Offset(0, (i - 1) * 30).Address
            suppoffset = (MbrNMBR + 8) * (i - 1)
            ConvBKMRK.Offset(suppoffset + 1, 8).Formula = "=CONCATENATE(IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$P$5:$P$99999=$B12)*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!Z$5:Z$99999)),DOLLAR(SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$P$5:$P$99999=$B12)*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!AJ$5:AJ$99999)),0)),"" of "",IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUMIF('Line Item Data'!$P:$P,$B12,'Line Item Data'!$Z:$Z),DOLLAR(SUMIF('Line Item Data'!$P:$P,$B12,'Line Item Data'!$AJ:$AJ),0)))"
            ConvBKMRK.Offset(suppoffset + 1, 8).AutoFill Destination:=Range(ConvBKMRK.Offset(suppoffset + 1, 8), ConvBKMRK.Offset(suppoffset + MbrNMBR, 8))
            ConvBKMRK.Offset(suppoffset + MbrNMBR + 1, 8).Formula = "=CONCATENATE(IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!Z$5:Z$99999)),DOLLAR(SUMPRODUCT(('Line Item Data'!$AI$5:$AI$99999<>""X"")*('Line Item Data'!$X$5:$X$99999<>'Line Item Data'!" & SuppCol & ")*('Line Item Data'!" & SuppCol & "<>""-"")*('Line Item Data'!AJ$5:AJ$99999)),0)),"" of "",IF(" & ConvBKMRK.Offset(suppoffset - 1, 8).Address & "=FALSE,SUM('Line Item Data'!$Z:$Z),DOLLAR(SUM('Line Item Data'!$AJ:$AJ),0)))"
        Next

        Sheets("Line item data").Range("Z5").Formula = "=IF(SUMPRODUCT(($X$5:$X5=$X5)*($P$5:$P5=$P5))>1,0,1)"
        'Sheets("Line item data").Range("Z5").AutoFill Destination:=Range(Sheets("Line item data").Range("Z5"), Sheets("Line item data").Range("A4").End(xlDown).Offset(0, 25))

End Sub

Sub modify_EnvironVar()

Set objWshShell = wscript.CreateObject("WScript.Shell")
'
' Run through the environment variables
'
strVariables = ""
For Each objEnvVar In objWshShell.Environment("Process")
    objEnvVar = strVariables & objEnvVar & Chr(13)
Next

'WshShell.Environment("Process")("[name of environ var]")


End Sub
Sub update_mbrs()
'
'For i = 60 To 208
'    Range(Range("H8").Offset(i, 0), Range("H8").Offset(i, 3)).Merge
'Next

For i = 60 To 208
    ActiveSheet.CheckBoxes("Check Box " & i).Value = True
Next


End Sub

Sub fibTemp()


Call FibonachiGenerator(0.01, 50, 1.25, ActiveCell)


End Sub

Sub FibonachiGenerator(StrtVal As Variant, TimesRpt As Integer, IncVal As Variant, RetStrt As Range)

RetStrt = StrtVal

For i = 1 To TimesRpt
    RetStrt.Offset(i, 0).Value = RetStrt.Offset(i - 1, 0).Value * IncVal
Next


End Sub
Sub TestFuzzy()

Set testwb = ActiveWorkbook
Set stdmfgwb = Workbooks("Standard Manufacturer Names.xlsx")
stdmfgwb.Activate

        For Each c In Range(Range("A2"), Range("B1").End(xlDown))
            Debug.Print c.Row
            On Error Resume Next
            c.Offset(0, 5).Value = testwb.ActiveSheet.Range("A1:O10").Find(what:=c.Value, lookat:=xlPart).Address
NOmfg:  Next

'testwb.Sheets("Ethicon").Range("P1").End(xlDown).Offset(1, 0).Value
'if any returned cells are empty or duplicate then remove
'find mode row
'use returned cells only in that row
'find their

Exit Sub
'::::::::::::::::::::::::::::::::::
errhndlNOMFG:
Resume NOmfg

End Sub
Sub testfuzzyAlt()

Set testwb = ActiveWorkbook
Set stdmfgwb = Workbooks("Standard Manufacturer Names.xlsx")
stdmfgwb.Activate

'sort by col B (only once per time it is run)

        Range("A2").Select
        strt = Now
        Do
            'if Application.CountIf(Range("B:B"), ActiveCell.Offset(0, 1).Value) =1 then
                
            'Else
            Debug.Print ActiveCell.Row
            On Error GoTo errhndlNOstd
            ActiveCell.Offset(0, 5).Value = testwb.ActiveSheet.Range("A1:O10").Find(what:=ActiveCell.Offset(0, 1).Value, lookat:=xlPart).Address
            GoTo NXTstd
NOstd:      For Each c In Range(ActiveCell, ActiveCell.Offset(Application.CountIf(Range("B:B"), ActiveCell.Offset(0, 1).Value) - 1, 0))
                On Error GoTo errhndlNOMFG
                c.Offset(0, 5).Value = testwb.ActiveSheet.Range("A1:O10").Find(what:=c.Value, lookat:=xlPart).Address
                c.Offset(Application.CountIf(Range("B:B"), c.Offset(0, 1).Value) - Application.CountIf(Range(c.Offset(0, 1), c.Offset(100, 1)), c.Offset(0, 1).Value), 0).Select
                GoTo NXTmfg
NOmfg:      Next
            'End If
NXTstd:     ActiveCell.Offset(Application.CountIf(Range("B:B"), ActiveCell.Offset(0, 1).Value), 0).Select
NXTmfg: Loop Until IsEmpty(ActiveCell.Offset(1, 0))
        ed = Now
'testwb.Sheets("Ethicon").Range("P1").End(xlDown).Offset(1, 0).Value
'if any returned cells are duplicate then fuzzy lookup to see which characters/words match more
'find mode row
'use returned cells only in that row
'find their

Exit Sub
'::::::::::::::::::::::::::::::::::
errhndlNOstd:
Resume NOstd

errhndlNOMFG:
Resume NOmfg

End Sub
