Option Explicit

Dim boCMS, boName, boPwd, boUrl, boUrlSL, boToken, url  As String


'-----------------------------------------------------------------------------------------------------------------------------------------
' spécial : UTF8
' trouvé sur http://forum.hardware.fr/hfr/Programmation/VB-VBA-VBS/code-conversion-ansi-sujet_79551_1.htm
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function Encode_UTF8(astr)
    Dim c
    Dim n
    Dim utftext
     
    utftext = ""
    n = 1
    Do While n <= Len(astr)
        c = AscW(Mid(astr, n, 1))
        If c < 128 Then
            utftext = utftext + Chr(c)
        ElseIf ((c >= 128) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        ElseIf ((c >= 2048) And (c < 65536)) Then
            utftext = utftext + Chr(((c \ 4096) Or 224))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else ' c >= 65536
            utftext = utftext + Chr(((c \ 262144) Or 240))
            utftext = utftext + Chr(((((c \ 4096) And 63)) Or 128))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
        n = n + 1
    Loop
    Encode_UTF8 = utftext
End Function
 
'   Char. number range  |        UTF-8 octet sequence
'      (hexadecimal)    |              (binary)
'   --------------------+---------------------------------------------
'   0000 0000-0000 007F | 0xxxxxxx
'   0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'   0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'   0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
Public Function Decode_UTF8(astr)
    Dim c0, c1, c2, c3
    Dim n
    Dim unitext
     
    If isUTF8(astr) = False Then
        Decode_UTF8 = astr
        Exit Function
    End If
     
    unitext = ""
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
         
        If (c0 And 240) = 240 And (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 240) * 65536 + (c1 - 128) * 4096) + (c2 - 128) * 64 + (c3 - 128)
            n = n + 4
        ElseIf (c0 And 224) = 224 And (c1 And 128) = 128 And (c2 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 224) * 4096 + (c1 - 128) * 64 + (c2 - 128))
            n = n + 3
        ElseIf (c0 And 192) = 192 And (c1 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 192) * 64 + (c1 - 128))
            n = n + 2
        ElseIf (c0 And 128) = 128 Then
            unitext = unitext + ChrW(c0 And 127)
            n = n + 1
        Else ' c0 < 128
            unitext = unitext + ChrW(c0)
            n = n + 1
        End If
    Loop
 
    Decode_UTF8 = unitext
End Function
 
'   Char. number range  |        UTF-8 octet sequence
'      (hexadecimal)    |              (binary)
'   --------------------+---------------------------------------------
'   0000 0000-0000 007F | 0xxxxxxx
'   0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'   0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'   0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
Public Function isUTF8(astr)
    Dim c0, c1, c2, c3
    Dim n
     
    isUTF8 = True
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
         
        If (c0 And 240) = 240 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
                n = n + 4
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 224) = 224 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 Then
                n = n + 3
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 192) = 192 Then
            If (c1 And 128) = 128 Then
                n = n + 2
            Else
                isUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 128) = 0 Then
            n = n + 1
        Else
            isUTF8 = False
            Exit Function
        End If
    Loop
End Function



'-----------------------------------------------------------------------------------------------------------------------------------------
' les fonctions utilisées un peu partout
'-----------------------------------------------------------------------------------------------------------------------------------------

Private Function getAttribute(o As WinHttp.WinHttpRequest, name As String)
    Dim tmpXML As MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    
    Set tmpXML = CreateObject("Microsoft.XMLDOM")
    tmpXML.LoadXML (o.ResponseText)
    
    'recup attribut
    For Each node In tmpXML.SelectNodes("//attrs/attr")
        If node.Attributes.getNamedItem("name").Text = name Then getAttribute = node.Text
    Next

End Function

' récupère le code erreur de l'API REST
Private Function getErrorCodeREST(o As WinHttp.WinHttpRequest)
    Dim tmpXML As MSXML2.DOMDocument
    
    If o.Status <> "200" Then
        Set tmpXML = CreateObject("Microsoft.XMLDOM")
        tmpXML.LoadXML (o.ResponseText)
        getErrorCodeREST = tmpXML.SelectSingleNode("/error/error_code").Text
    Else
        getErrorCodeREST = ""
    End If
End Function

' affiche l'erreur de l'API REST
Private Sub afficheErrorREST(o As WinHttp.WinHttpRequest, titre As String, msg As String)
    Dim tmpXML As MSXML2.DOMDocument
    Dim errorCode, errorMsg As String
        
    Set tmpXML = CreateObject("Microsoft.XMLDOM")
    tmpXML.LoadXML (o.ResponseText)
    errorCode = tmpXML.SelectSingleNode("/error/error_code").Text
    errorMsg = tmpXML.SelectSingleNode("/error/message").Text
    Call MsgBox(msg & vbCrLf & _
        "------------------------------------------------------" & vbCrLf & _
        "détails : " & errorCode & vbCrLf & errorMsg, _
        vbCritical, "Erreur " & titre)
    'todo : en fonction du code erreur, mettre un message explicite
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
'les fonctions propres à BO, avec une url spécifique
'-----------------------------------------------------------------------------------------------------------------------------------------

' récupère le nom du dossier à partir de l'id
Private Function getFolder(id)
    Dim tmpHTTP As WinHttp.WinHttpRequest
    Dim tmpXML As MSXML2.DOMDocument
    Dim s As String

    url = boUrl & "/infostore/" & id
    Debug.Print url
    
    Set tmpHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    tmpHTTP.Open "GET", url, False
    tmpHTTP.SetRequestHeader "Accept", "application/xml"
    tmpHTTP.SetRequestHeader "X-SAP-LogonToken", boToken
    tmpHTTP.Send ""
        
    If tmpHTTP.Status <> "200" Then
        Call afficheErrorREST(tmpHTTP, "getFolder", "Error gettin folder " & id)
    Else
        s = getAttribute(tmpHTTP, "name")
        getFolder = Decode_UTF8(s)
    End If
    
    'todo : récupérer l'arborescence complète
    
End Function

Private Sub logoff()
    Dim tmpHTTP As WinHttp.WinHttpRequest
    Dim url As String

    url = boUrl & "/logoff"
    Debug.Print url
    Debug.Print boToken
    
    Set tmpHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    tmpHTTP.Open "POST", url, False
    tmpHTTP.SetRequestHeader "Accept", "application/xml"
    tmpHTTP.SetRequestHeader "X-SAP-LogonToken", boToken
    tmpHTTP.Send ""
    If tmpHTTP.Status <> "200" Then
        Call afficheErrorREST(tmpHTTP, "logoff", "Error on logout, check in CMC if session is always active")
    End If
End Sub

' recup id de login, se connecte et recupere le token
Private Sub logon()
    Dim tmpHTTP As WinHttp.WinHttpRequest
    Dim tmpXML As MSXML2.DOMDocument

    ' recup paramétrage
    boCMS = Sheets("config").Cells(1, 2).Value
    boName = Sheets("config").Cells(2, 2).Value
    boPwd = Sheets("config").Cells(3, 2).Value
    boUrl = Sheets("config").Cells(4, 2).Value
    boUrlSL = boUrl & "/sl/v1"
    
    'recup du modele XML de logon
    Set tmpHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    tmpHTTP.Open "GET", boUrl & "/logon/long", False
    tmpHTTP.SetRequestHeader "Content-type", "application/xml"
    tmpHTTP.SetRequestHeader "Accept", "application/xml"
    On Error GoTo pbHttp
    tmpHTTP.Send ""
    
    If tmpHTTP.Status <> "200" Then
        Call afficheErrorREST(tmpHTTP, "logon", "Error on logon request, check server name, port, url and WACS running")
        End
    End If

    Set tmpXML = CreateObject("Microsoft.XMLDOM")
    tmpXML.LoadXML (tmpHTTP.ResponseText)
    
    'modif des attributs
    tmpXML.SelectSingleNode("/attrs/attr[0]").Text = boName
    tmpXML.SelectSingleNode("/attrs/attr[1]").Text = boPwd
'    Debug.Print objXML.XML
        
    'post
    tmpHTTP.Open "POST", boUrl & "/logon/long", False
    tmpHTTP.SetRequestHeader "Content-type", "application/xml"
    tmpHTTP.SetRequestHeader "Accept", "application/xml"
    tmpHTTP.Send (tmpXML.XML)
    
    If tmpHTTP.Status <> "200" Then
        Call afficheErrorREST(tmpHTTP, "logon", "Error on logon, check username and password")
        End
    End If
'    objXML.LoadXML (objHTTP.ResponseText)
    boToken = tmpHTTP.GetResponseHeader("X-SAP-LogonToken")
    Debug.Print "token=" & boToken
    
    Exit Sub
    
pbHttp:
    Call MsgBox("Error on logon, can't view BI4, check server name !" & vbCrLf & "try ping " & boCMS, vbCritical, "logon")
    End
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------
' les macros publiques
'-----------------------------------------------------------------------------------------------------------------------------------------


Public Sub refreshListUnivers()
    Dim objHTTP As WinHttp.WinHttpRequest
    Dim objXML As MSXML2.DOMDocument
    Dim oNodeXML, oSubNodeXML As MSXML2.IXMLDOMNode
    Dim folderId, errorCodeREST, t As String
    Dim i, l As Integer
        
    Call logon
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set objXML = CreateObject("Microsoft.XMLDOM")
    
    'efface onglet univers, sauf entete
    Sheets("liste univers").Range("A2:Z65000").ClearContents
    
    'recup univers
    ' en plusieurs passes
    l = 0
    Do
        url = boUrlSL & "/universes?offset=" & l & "&limit=50"
        Debug.Print url
        objHTTP.Open "GET", url, False
        objHTTP.SetRequestHeader "Content-type", "application/xml"
        objHTTP.SetRequestHeader "Accept", "application/xml"
        objHTTP.SetRequestHeader "X-SAP-LogonToken", boToken
        objHTTP.Send ""
        
        errorCodeREST = getErrorCodeREST(objHTTP)
        If errorCodeREST <> "" And errorCodeREST <> "WSR 00400" Then
            Call afficheErrorREST(objHTTP, "RefreshlistUnivers", "Error getting list of universes")
            Call logoff
            End
        End If
        
        
        objXML.LoadXML (objHTTP.ResponseText)
        'Debug.Print objHTTP.ResponseText
        i = 0
        For Each oNodeXML In objXML.SelectNodes("/universes/universe")
            t = "/universes/universe[" & i & "]"
            Sheets("liste univers").Cells(l + 2, 1) = objXML.SelectSingleNode(t & "/id").Text
            Sheets("liste univers").Cells(l + 2, 2) = objXML.SelectSingleNode(t & "/cuid").Text
            Sheets("liste univers").Cells(l + 2, 4) = objXML.SelectSingleNode(t & "/type").Text
            folderId = objXML.SelectSingleNode(t & "/folderId").Text
            Sheets("liste univers").Cells(l + 2, 3) = getFolder(folderId)
            Sheets("liste univers").Cells(l + 2, 5) = Decode_UTF8(objXML.SelectSingleNode(t & "/name").Text)
            
            l = l + 1
            i = i + 1
        Next
    Loop While i > 0
        
logoff:
    Call logoff
End Sub
