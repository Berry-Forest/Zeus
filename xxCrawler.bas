Attribute VB_Name = "xxCrawler"
Sub UOMcrawler_CIA(ProdNmbr, ReturnStrt)

'Dim xhrRequest As ServerXMLHTTP60 'leaves the connection open
Dim xhrRequest As XMLHTTP60
Dim crawlerUOM As String
Dim pkgUOM As String
Dim PkgArray(1 To 18) As String
PkgArray(1) = "/bx"
PkgArray(2) = "/ca"
PkgArray(3) = "/cs"
PkgArray(4) = "/case"
PkgArray(5) = "/box"
PkgArray(6) = "/pk"
PkgArray(7) = "/pack"
PkgArray(8) = "/crtn"
PkgArray(9) = "/carton"
PkgArray(10) = "per box"
PkgArray(11) = "per case"
PkgArray(12) = "per pack"
PkgArray(13) = "/rl"
PkgArray(14) = "/roll"
PkgArray(15) = "per roll"
PkgArray(16) = "/bg"
PkgArray(17) = "/bag"
PkgArray(18) = "per bag"

'~~~~~~~~~~~~~~~~~~~~~
'UOMref = Range("PA3").Address
'ProdNmbr = "15120"
'~~~~~~~~~~~~~~~~~~~~~
'Range(ReturnStrt, Range("ZA1").End(xlDown)).ClearContents
'Range("ZA1:ZB2").Value = "x"

urlLnk = "https://www.ciamedical.com/search/" & ProdNmbr
UOMfndFLG = 0

On Error GoTo ERR_NoConnection
Set xhrRequest = New XMLHTTP60
xhrRequest.Open "GET", urlLnk, False
xhrRequest.send
CIAtxt = xhrRequest.responseText
On Error GoTo 0

'If Not InStr(CIAtxt, "There is no product that matches the search criteria.") > 0 Then Exit Sub

Do

'find the next item with CASE/BOX/PACK
'--------------------------------------------------
srchstr = ">per BOX</span>"
cropmark = InStrRev(CIAtxt, srchstr) - 20

srchstr2 = ">per CASE</span>"
cropmark2 = InStrRev(CIAtxt, srchstr2) - 20
If cropmark2 > cropmark Then cropmark = cropmark2

srchstr3 = ">per PACK</span>"
cropmark3 = InStrRev(CIAtxt, srchstr3) - 20
If cropmark3 > cropmark Then cropmark = cropmark3

srchstr4 = ">per ROLL</span>"
cropmark4 = InStrRev(CIAtxt, srchstr4) - 20
If cropmark4 > cropmark Then cropmark = cropmark4

srchstr5 = ">per BAG</span>"
cropmark5 = InStrRev(CIAtxt, srchstr5) - 20
If cropmark5 > cropmark Then cropmark = cropmark5

'from cropmark back to href srch for each pkg qty
'--------------------------------------------------
If cropmark < 0 Then GoTo EndSrch
CIAtxt = Left(CIAtxt, cropmark)
'Debug.Print CIAtxt
UOMsrchstr = LCase(Mid(CIAtxt, InStrRev(CIAtxt, "<a href=")))
'Debug.Print UOMsrchStr

'find the UOM in the new srchStr
'--------------------------------------------------
'look also for pk, cs, crtn, carton case, pack, box, and per for each instead of /

OrigDesc = "[Desc]" & Mid(UOMsrchstr, InStr(UOMsrchstr, ">") + 1, InStr(UOMsrchstr, "</a><span style=") - InStr(UOMsrchstr, ">") - 1)
ttlUOM = ""
For i = 1 To 18
    pkgUOM = 0
    pkgvar = PkgArray(i)
    If InStr(UOMsrchstr, pkgvar) > 0 Then
        
        'Check for combo pkg
        '-------------------------
        For II = 1 To 18
            pkgvar2 = PkgArray(II)
            If InStr(UOMsrchstr, Mid(pkgvar, 2) & pkgvar2) > 0 Then
                UOMstr = " " & Trim(Mid(UOMsrchstr, InStr(UOMsrchstr, Mid(pkgvar, 2) & pkgvar2) - 5, 5))
                pkgUOM = Mid(UOMstr, InStrRev(UOMstr, " ") + 1)
                pkgUOM = FUN_NumberOnly(pkgUOM)
                ttlUOM = ttlUOM & " " & pkgUOM & Mid(pkgvar, 2) & pkgvar2
                If Not Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-1, 0).End(xlDown)), pkgUOM) > 0 Then ReturnStrt.Offset(-1, 0).End(xlDown).Offset(1, 0).Value = pkgUOM
                UOMsrchstr = Replace(UOMsrchstr, Mid(pkgvar, 2) & pkgvar2, "")
            End If
        Next
        
        'Check for isolated pkg
        '-------------------------
        If InStr(UOMsrchstr, pkgvar) > 0 Then
            UOMstr = Left(UOMsrchstr, InStr(UOMsrchstr, pkgvar))
            If IsNumeric(Mid(UOMstr, InStrRev(UOMstr, " ") + 1, 1)) Then
                crawlerUOM = " " & Trim(Mid(UOMstr, InStrRev(UOMstr, " ") + 1, 4))
            Else
                UOMstr = Left(UOMstr, InStrRev(UOMstr, " ") - 1)
                crawlerUOM = " " & Trim(Mid(UOMstr, InStrRev(UOMstr, " ") + 1, 4))
            End If
            'UOMstr = " " & Trim(Mid(UOMsrchstr, InStr(UOMsrchstr, pkgvar) - 5, 5))
            'crawlerUOM = Mid(UOMstr, InStrRev(UOMstr, " ") + 1)
            crawlerUOM = FUN_NumberOnly(crawlerUOM)
            If pkgUOM > 0 Then
                ttlUOM = pkgUOM * crawlerUOM & "/" & "ea" & " " & crawlerUOM & pkgvar & " " & ttlUOM
                If Not Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-1, 0).End(xlDown)), pkgUOM * crawlerUOM) > 0 Then ReturnStrt.Offset(-1, 0).End(xlDown).Offset(1, 0).Value = pkgUOM * crawlerUOM
                If Not Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-1, 0).End(xlDown)), crawlerUOM) > 0 Then ReturnStrt.Offset(-1, 0).End(xlDown).Offset(1, 0).Value = crawlerUOM
            Else
                ttlUOM = crawlerUOM & pkgvar & " " & ttlUOM
                If Not Application.CountIf(Range(ReturnStrt, ReturnStrt.Offset(-1, 0).End(xlDown)), crawlerUOM) > 0 Then ReturnStrt.Offset(-1, 0).End(xlDown).Offset(1, 0).Value = crawlerUOM
            End If
        End If
    End If
Next
If Not ttlUOM = "" Then
    If Not Application.CountIf(Range(UOMref, UOMref.Offset(0, -1).End(xlToRight)), Trim(Replace(ttlUOM, "  ", " "))) > 0 Then
        UOMfndFLG = 1
        UOMref.Offset(0, -1).End(xlToRight).Offset(0, 1).Value = Trim(Replace(ttlUOM, "  ", " "))
        UOMref.Offset(0, -1).End(xlToRight).Offset(0, 1).Value = OrigDesc
    End If
End If

nxtPkg:
CIAtxt = Left(CIAtxt, InStrRev(CIAtxt, "<a class=""s_thumb"))
'Debug.Print CIAtxt

Loop Until cropmark < 0

EndSrch:
'If UOMfndFLG = 1 Then
'    UOMref.Offset(0, -1).End(xlToRight).Offset(0, 1).Value = urlLnk
'Else
'    UOMref.Offset(0, -1).End(xlToRight).Offset(0, 1).Value = "no CIA found"
'End If

Exit Sub
':::::::::::::::::::::::::::::::::::::::
ERR_NoConnection:
Exit Sub



End Sub

Sub testResponse()

Dim aRequest
Set aRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
aRequest.Open "GET", "http://www.google.com", False
aRequest.send
Debug.Print aRequest.responseText
wscript.Echo aRequest.responseText

End Sub
Sub Getsource()


Dim URL, resp As String
Dim req As New WinHttpRequest

URL = "https://www.ciamedical.com/search/F145"
req.Open "GET", URL, False
req.send
Sheets("Sheet2").Range("A1").Value = req.responseText
resp = req.responseText

End Sub
Sub FTPeinstein()

'Const sUrl As String = "https://www.marketplaceprocure.com/contracts/detailContracts.html?R=1912277_29832&Tab=2"


Dim oRequest As WinHttp.WinHttpRequest
Dim sResult As String
 
'On Error GoTo Err_DoSomeJob
 
Set oRequest = New WinHttp.WinHttpRequest
With oRequest
'
'    .Open "GET", "https://sso2.alliancewebs.net/oamfed/idp/samlv20?SAMLRequest=fZLdb4IwFMX%2fFdJ3AfnQ0KAJgmYmbiPC9rC3itfZpLSst6j774e4D80SX2%2fP7%2fackxsjq0VDk9bs5Ro%2bWkBjnWohkfYPE9JqSRVDjlSyGpCaihbJ44p6tksbrYyqlCBXyH2CIYI2XEliLbMJKbwojWbjaJR6Q38WJlnkjefucDROF24QpcHQC8aLeZglIbFeQWMHTki3p6MRW1hKNEyabuQOw4EbDHy39Hzq%2bzT0386avPuPH2BCdkwgEGuhdAV91t9R1iXmkpl%2b9d6YBqnjICrPZkJwJis4wgZtCcZRrN7B1uHbxjknPXidj%2fy7ghmXWy7f76ffXERIH8oyH%2bTPRUms5KeRVElsa9AF6AOv4GW9uvHz387Zw57JrQBtM2xOZBqfR7SvRk%2fvsbFzrYwvN%2fDU%2bV1muRK8%2brQSIdQx1cDMX3vO9MLdnsv0Cw%3d%3d&RelayState=login.aspx%3fBCM%3dhttp%253a%252f%252fwww.marketplaceprocure.com%253a80%252fsecurity%252fauth%26ReturnUrl%3dhttp%253a%252f%252fwww.marketplaceprocure.com%253a80%252fsecurity%252fauth", True
'    .SetCredentials "bforrest", "psych101!", HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
'    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
'    .send "{range:9129370}"
'    .WaitForResponse
'    sResult = .responseText
    
    .Open "GET", "https://www.medline.com/product/Quincke-Spinal-Needles/Epidural-Needles/Z05-PF19762", True
    '.SetCredentials "bforrest", "psych101!", 0
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    .send "{range:9129370}"
    .WaitForResponse
    sResult = .responseText

End With

Range("A1").Value = sResult
'ActiveSheet.Paste

 
'Exit_DoSomeJob:
'    On Error Resume Next
'    Set oRequest = Nothing
'    Exit Sub
'
'Err_DoSomeJob:
'    MsgBox Err.Description, vbExclamation, Err.Number
'    Resume Exit_DoSomeJob
'
End Sub
Sub ftp2()

Dim xhrRequest As XMLHTTP60

'urllnk = "http://catalog.bd.com/nexus-ecat/getProductDetail?productId=405234&parentCategory=&parentCategoryName=&categoryId=1220&categoryName=Anesthesia%20Needles,%20Syringes,%20Trays&searchUrl="
urlLnk = "http://ecatalog.baxter.com/ecatalog/loadproductsearchresults.html?cid=20016&lid=10001&hid=20001&searchCriteriaPattern=934070"

Set xhrRequest = New XMLHTTP60
xhrRequest.Open "GET", urlLnk, False
xhrRequest.send

Debug.Print Mid(xhrRequest.responseText, 1, 5000)


End Sub
Sub WebReq()

link = "https://www.marketplaceprocure.com/contracts/detailContracts.html?R=1912277_29832&Tab=2" & str(Rnd())
Set htm = CreateObject("htmlFile")
Dim objHttp

    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
    objHttp.Open "GET", link, False

    objHttp.send
    htm.body.innerHTML = objHttp.responseText
    Set objHttp = Nothing
End Sub
Function URLGet(URL)

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", URL, False
    http.send
    URLGet = http.responseText
    
End Function
Sub UOMcrawler_Meta()

Dim xhrRequest As XMLHTTP60

'to start with maybe a google search to find which metasrchsare present then hit each in order of priority, instead of having to test each one

ProdNmbr = "5102710"
SuppNm = "bard"

googleurl = "https://www.google.com/search?q=" & ProdNmbr & "+" & SuppNm
Set xhrRequest = New XMLHTTP60
xhrRequest.Open "GET", googleurl, False
xhrRequest.send
MetaRspn = xhrRequest.responseText

'Debug.Print xhrRequest.responseText
If InStr(LCase(MetaRspn), "https://www.ciamedical.com") > 0 Then
    'call UOMcrawler_CIA
End If
If Not UOMfndFLG = 1 And InStr(LCase(MetaRspn), "https://www.medicalsupplydepot.com") > 0 Then
    'call UOMcrawler_CIA
End If
If Not UOMfndFLG = 1 And InStr(LCase(MetaRspn), "https://www.berktree.com") > 0 Then
    'call UOMcrawler_CIA
End If
If Not UOMfndFLG = 1 And InStr(LCase(MetaRspn), "https://www.beesmed.com") > 0 Then
    'call UOMcrawler_CIA
End If
If Not UOMfndFLG = 1 And InStr(LCase(MetaRspn), "https://www.aaawholesalesompany.com") > 0 Then
    'call UOMcrawler_CIA
End If
If Not UOMfndFLG = 1 And InStr(LCase(MetaRspn), "https://www.esutures.com") > 0 Then
    'call UOMcrawler_CIA
End If
If Not UOMfndFLG = 1 And InStr(LCase(MetaRspn), "https://www.amazon.com") > 0 Then
    'call UOMcrawler_CIA
End If


End Sub
