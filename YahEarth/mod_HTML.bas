Attribute VB_Name = "mod_HTML"
Option Explicit

Public strBuffer As String

Public Function ProcessRoom(strRoom As String, strTopic As String, WB As WebBrowser, Optional intUsers As String)
    strBuffer = strBuffer & "<br><b><font color=" & Chr(34) & "#008000" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strRoom & "</b></font><font color=" & Chr(34) & "#000000" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & "> (" & strTopic & ") [Users: " & intUsers & "]</font></font><br><br>"
    WB.Document.write "<br><b><font color=" & Chr(34) & "#008000" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strRoom & "</b></font><font color=" & Chr(34) & "#000000" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & "> (" & strTopic & ") [Users: " & intUsers & "]</font></font><br><br>"
    WB.Document.parentwindow.Scroll 0, 999999999
End Function

Public Function ProcessError(WB As WebBrowser, strError As String, Optional blWarning As Boolean = False)
    Dim strUser As String
    If blWarning = True Then strUser = "Warning" Else strUser = "Error"
    strBuffer = strBuffer & "<b><font color=" & Chr(34) & "#FF0000" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strUser & ":</b><font color=" & Chr(34) & "#000000" & Chr(34) & "> </b>" & strError & "<br></font></font>"
    WB.Document.write "<b><font color=" & Chr(34) & "#FF0000" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strUser & ":</b><font color=" & Chr(34) & "#000000" & Chr(34) & "> </b>" & strError & "<br></font></font>"
    WB.Document.parentwindow.Scroll 0, 999999999
End Function

Public Function ProcessRoomUser(strUser As String, WB As WebBrowser, strState As String)
    strBuffer = strBuffer & "<b><font color=" & Chr(34) & "#900000" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strUser & "</b><font color=" & Chr(34) & "#000000" & Chr(34) & "> </b>" & strState & "<br></font></font>"
    WB.Document.write "<b><font color=" & Chr(34) & "#900000" & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strUser & "</b><font color=" & Chr(34) & "#000000" & Chr(34) & "> </b>" & strState & "<br></font></font>"
    WB.Document.parentwindow.Scroll 0, 999999999
End Function

Public Function ProcessStatus(strUser As String, WB As WebBrowser, blOnline As Boolean)
    Dim strCode As String
    
    strCode = "<table height=" & Tag(40) & " border=" & Tag(0) & " width=" & Tag("100%") & ">" & vbCrLf & "<tr>" & vbCrLf & "<td>" & vbCrLf
    If blOnline = False Then
        strCode = strCode & "<font face=" & Tag("Arial") & " size=" & Tag("2") & "><font color=" & Tag("#820000") & "><img src=" & Tag("file://" & App.Path & "\Resources\Pictures\Icons\offline.gif") & " border=0>&nbsp;<b>" & strUser & "</b></font> <font color=" & Tag("#000000") & ">is now Offline<br></font></font>" & vbCrLf
    Else
        strCode = strCode & "<font face=" & Tag("Arial") & " size=" & Tag("2") & "><font color=" & Tag("#820000") & "><img src=" & Tag("file://" & App.Path & "\Resources\Pictures\Icons\online.gif") & " border=0>&nbsp;<b>" & strUser & "</b></font> <font color=" & Tag("#000000") & ">is now Online<br></font></font>" & vbCrLf
    End If
    strCode = strCode & "</td>" & vbCrLf & "</tr>" & vbCrLf & "</table>"
    
    strBuffer = strBuffer & strCode
    WB.Document.write strCode
    WB.Document.parentwindow.Scroll 0, 999999999
End Function

Public Function ProcessHTML(strUser As String, ByVal strText As String, WB As WebBrowser, Optional blPM As Boolean = False, Optional blProcessOnly As Boolean = False) As String
    Dim Allow() As Variant
    Dim X As Integer
    Dim Color As String
    Dim Scroll As Boolean
    Dim strStart As String
    Dim strEnd As String
    
    Scroll = True
    'Scroll is on ToDo list
    
    Allow = Array("font", "b>", "i>", "u>", "/b>", "/i>", "/u>", "/font>", "alt", "fade", "/fade>", "/alt>", "#")
    
    strText = ReplaceHexColors(strText)
    
    strText = Replace(strText, "<black>", "<font color=" & Chr(34) & "#000000" & Chr(34) & ">", , , vbTextCompare)
    strText = Replace(strText, "<green>", "<font color=" & Chr(34) & "#008000" & Chr(34) & ">", , , vbTextCompare)
    strText = Replace(strText, "<blue>", "<font color=" & Chr(34) & "#0000FF" & Chr(34) & ">", , , vbTextCompare)
    strText = Replace(strText, "<red>", "<font color=" & Chr(34) & "#FF0000" & Chr(34) & ">", , , vbTextCompare)
    strText = Replace(strText, "</red>", "</font>", , , vbTextCompare)
    strText = Replace(strText, "</green>", "</font>", , , vbTextCompare)
    strText = Replace(strText, "</blue>", "</font>", , , vbTextCompare)
    strText = Replace(strText, "</black>", "</font>", , , vbTextCompare)
    
    strText = Replace(strText, "<", "&lt;")
    For X = LBound(Allow) To UBound(Allow)
        If Options.blDisableFontStyle = True Then
            If Not Allow(X) = "font" And Not Allow(X) = "fade" And Not Allow(X) = "alt" And Not Allow(X) = "#" Then
                strText = Replace(strText, "&lt;" & Allow(X), "", , , vbTextCompare)
            Else
                strText = Replace(strText, "&lt;" & Allow(X), "<" & Allow(X), , , vbTextCompare)
            End If
        Else
            strText = Replace(strText, "&lt;" & Allow(X), "<" & Allow(X), , , vbTextCompare)
        End If
    Next X
        
    strText = ReplaceTagColors(strText)
    strText = ReplaceSize(strText)
    strText = ProcessSmileys(strText)
    strText = DoUrls(strText)
    
    If strUser = YMSG.strUser Then
        Color = "#000000"
        strStart = ""
        strEnd = ""
    Else
        Color = "#0000FF"
    End If
    
    strText = AddEndTags(strText)
    
    If Not blProcessOnly = True Then
        If blPM = True Then
            WB.Document.write "<font color=" & Chr(34) & Color & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & "><b>" & strStart & strUser & strEnd & ": </b><font color=" & Chr(34) & "#000000" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strText & "<br>"
        Else
            strBuffer = strBuffer & "<font color=" & Chr(34) & Color & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & "><b>" & strStart & "<a href=" & Chr(34) & "yahearth:" & strUser & Chr(34) & "style=" & Chr(34) & "color: " & Color & "; visited: #0000FF; active: #0000FF; text-decoration: none;" & Chr(34) & ">" & strUser & "</a>" & strEnd & ": </b><font color=" & Chr(34) & "#000000" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strText & "<br>"
            WB.Document.write "<font color=" & Chr(34) & Color & Chr(34) & " face=" & Chr(34) & "Arial" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & "><b>" & strStart & "<a href=" & Chr(34) & "yahearth:" & strUser & Chr(34) & "style=" & Chr(34) & "color: " & Color & "; visited: #0000FF; active: #0000FF; text-decoration: none;" & Chr(34) & ">" & strUser & "</a>" & strEnd & ": </b><font color=" & Chr(34) & "#000000" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">" & strText & "<br>"
        End If
        
        If Scroll = True Then
            WB.Document.parentwindow.Scroll 0, 999999999
        End If
    End If
    
    Debug.Print strText
    
    ProcessHTML = strText
End Function


Public Function GenerateHTML(Buffer As RichTextBox) As String
    Dim strMsg As String
    Dim X As Integer
    Dim strLastFont As String
    Dim strLastSize As String
    Dim strLastColor As ColorConstants
    Dim Bold As Boolean
    Dim Italic As Boolean
    Dim Underline As Boolean
        
    strLastColor = vbBlack
       
    For X = 1 To Len(Buffer.Text)
        Buffer.SelStart = X
        'Bold
        If Buffer.SelBold = True Then
            If Bold = False Then
                strMsg = strMsg & "<b>"
                Bold = True
            End If
        Else
            If Bold = True Then
                strMsg = strMsg & "</b>"
                Bold = False
            End If
        End If
        
        'Italic
        If Buffer.SelItalic = True Then
            If Italic = False Then
                strMsg = strMsg & "<i>"
                Italic = True
            End If
        Else
            If Italic = True Then
                strMsg = strMsg & "</i>"
                Italic = False
            End If
        End If

        'Italic
        If Buffer.SelUnderline = True Then
            If Underline = False Then
                strMsg = strMsg & "<u>"
                Underline = True
            End If
        Else
            If Underline = True Then
                strMsg = strMsg & "</u>"
                Underline = False
            End If
        End If
        
        If Not Buffer.SelColor = strLastColor Then
            strMsg = strMsg & "[" & RGB2Hex(Buffer.SelColor) & "m"
            strLastColor = Buffer.SelColor
        End If
        
        'Font
        If Not Buffer.SelFontName = strLastFont Then
            strMsg = strMsg & "<font face=" & Chr(34) & Buffer.SelFontName & Chr(34) & ">"
            strLastFont = Buffer.SelFontName
        End If
        
        If Not Buffer.SelFontSize = strLastSize Then
            strMsg = strMsg & "<font size=" & Chr(34) & Buffer.SelFontSize & Chr(34) & ">"
            strLastSize = Buffer.SelFontSize
        End If
        
        strMsg = strMsg & Mid(Buffer.Text, X, 1)
    Next X
    
    GenerateHTML = strMsg
    Buffer.Text = ""
End Function

Public Function ReplaceTagColors(ByVal strText As String) As String
    Dim strSub As String
    Dim strLeft As String
    Dim strRight As String
    Dim strColor As String
    Dim i As Integer
    Dim X As Integer
    
    i = InStr(1, strText, "<#")
    Do While Not i = 0
        strLeft = Left(strText, i - 1)
        strSub = Mid(strText, i)
        X = InStr(strSub, ">")
        If X = 0 Then
            'I do that later
            strText = strLeft & strSub
            GoTo exError
        End If
        strRight = Mid(strSub, X + 1)
        strSub = Left(strText, i - 1)
        strColor = Mid(strSub, 2, 6)
        If Options.blDisableFontStyle = True Then
            strText = strLeft & strRight
        Else
            strText = strLeft & "<font color=" & Chr(34) & "#" & strColor & Chr(34) & ">" & strRight
        End If
        
exError:
        i = InStr(i + 1, strText, "<#")
        DoEvents
    Loop
    ReplaceTagColors = strText
End Function

Public Function ReplaceSize(ByVal strText As String) As String
    Dim strSub As String
    Dim strLeft As String
    Dim strRight As String
    Dim strFont As String
    
    Dim strFace As String
    Dim strColor As String
    Dim strSize As String
    
    Dim i As Integer
    Dim X As Integer
    
    i = InStr(strText, "<font")
    Do While Not i = 0
        strLeft = Left(strText, i - 1)
        strSub = Mid(strText, i)
        X = InStr(strSub, ">")
        If X = 0 Then
            If Left(LTrim(strSub), 1) = "<" Then
                strSub = "&lt;" & Mid(LTrim(strSub), 2)
            End If
            strText = strLeft & strSub
            GoTo exError
        End If
        strRight = Mid(strSub, X + 1)
        strSub = Left(strSub, X - 1)
        strSub = strSub & " >"
        If Not InStr(1, strSub, " inf ", vbTextCompare) = 0 Then
            strText = strLeft & " " & strRight
        Else
            strFont = "<font "
            If Not InStr(1, strSub, "size", vbTextCompare) = 0 Then
                strSize = Parse("size=", " ", strSub)
                strSize = Trim(Replace(strSize, Chr(34), ""))
                If Options.blDisableFontStyle = True Then strSize = "10"
                strFont = strFont & "style=" & Chr(34) & "font-size: " & strSize & "pt;" & Chr(34) & " "
            End If
            If Not InStr(1, strSub, "face", vbTextCompare) = 0 Then
                strFace = Parse("face=", " ", strSub)
                strFace = Trim(Replace(strFace, Chr(34), ""))
                If Options.blDisableFontStyle = True Then strFace = "Arial"
                strFont = strFont & "face=" & Chr(34) & strFace & Chr(34) & " "
            End If
            If Not InStr(1, strSub, "tattoo", vbTextCompare) = 0 Then
                strFont = strFont & "face=" & Chr(34) & "Webdings" & Chr(34) & " "
            End If
            If Not InStr(1, strSub, "color", vbTextCompare) = 0 Then
                strColor = Parse("color=", " ", strSub)
                strColor = Trim(Replace(strColor, Chr(34), ""))
                If Options.blDisableFontStyle = True Then strColor = "#000000"
                strFont = strFont & "color=" & Chr(34) & strColor & Chr(34) & " "
            End If
            strFont = strFont & ">"
            strText = strLeft & strFont & strRight
        End If
exError:
        i = InStr(i + 1, strText, "<font")
    Loop
    ReplaceSize = strText
End Function

Public Function ReplaceHexColors(ByVal strText As String) As String
    Dim i As Integer
    Dim X As Integer
    Dim strLeft As String
    Dim strRight As String
    Dim strSub As String
    Dim strColor As String
    Dim strCase As String
    Dim strCode As String
    
    i = InStr(strText, "[")
    Do While Not i = 0
        strLeft = Left(strText, i - 1)
        strSub = Mid(strText, i)
        X = InStr(strSub, "m")
        If X = 0 Then GoTo exError
        strRight = Mid(strSub, X + 1)
        strCase = Left(strSub, X - 1)
        strCase = Mid(Trim(strCase), 3)
        Select Case strCase
            Case 1: strCode = "<b>"
            Case 2: strCode = "<i>"
            Case 3: strCode = "<s>"
            Case 4: strCode = "<u>"
            Case "x1": strCode = "</b>"
            Case "x2": strCode = "</i>"
            Case "x3": strCode = "</s>"
            Case "x4": strCode = "</u>"
            Case 30: strCode = "<font color=" & Chr(34) & "#000000" & Chr(34) & ">"
            Case 31: strCode = "<font color=" & Chr(34) & "#0000FF" & Chr(34) & ">"
            Case 32: strCode = "<font color=" & Chr(34) & "#00FF00" & Chr(34) & ">"
            Case 33: strCode = "<font color=" & Chr(34) & "#848284" & Chr(34) & ">"
            Case 34: strCode = "<font color=" & Chr(34) & "#008200" & Chr(34) & ">"
            Case 35: strCode = "<font color=" & Chr(34) & "#FF0084" & Chr(34) & ">"
            Case 36: strCode = "<font color=" & Chr(34) & "#820082" & Chr(34) & ">"
            Case 37: strCode = "<font color=" & Chr(34) & "#FF8200" & Chr(34) & ">"
            Case 38: strCode = "<font color=" & Chr(34) & "#FF0000" & Chr(34) & ">"
            Case 39: strCode = "<font color=" & Chr(34) & "#848200" & Chr(34) & ">"
            Case Else
                If Left(strCase, 1) = "#" Then
                    strCode = "<font color=" & Chr(34) & strCase & Chr(34) & ">"
                End If
        End Select
        If Options.blDisableFontStyle = True Then
            strText = strLeft & " " & strRight
        Else
            strText = strLeft & strCode & strRight
        End If
exError:
        i = InStr(i + 1, strText, "[")
    Loop
    ReplaceHexColors = strText
    'Debug.Print strText
End Function

Public Function AddEndTags(ByVal strText As String) As String
    Dim i As Integer
    Dim X As Integer
    Dim C As Integer
    
    i = InStr(LCase(strText), "<b>")
    Do While Not i = 0
        If InStr(i + 1, LCase(strText), "</b>") = 0 Then C = C + 1
        i = InStr(i + 1, LCase(strText), "<b>")
    Loop
    For X = 1 To C
        strText = strText & "</b>"
    Next X
    
    i = InStr(LCase(strText), "<u>")
    Do While Not i = 0
        If InStr(i + 1, LCase(strText), "</u>") = 0 Then C = C + 1
        i = InStr(i + 1, LCase(strText), "<u>")
    Loop
    For X = 1 To C
        strText = strText & "</u>"
    Next X
    
    i = InStr(LCase(strText), "<i>")
    Do While Not i = 0
        If InStr(i + 1, LCase(strText), "</i>") = 0 Then C = C + 1
        i = InStr(i + 1, LCase(strText), "<i>")
    Loop
    For X = 1 To C
        strText = strText & "</i>"
    Next X
    
    i = InStr(LCase(strText), "<font")
    Do While Not i = 0
        If InStr(i + 1, LCase(strText), "</font>") = 0 Then C = C + 1
        i = InStr(i + 1, LCase(strText), "<font")
    Loop
    For X = 1 To C
        strText = strText & "</font>"
    Next X
    
    AddEndTags = strText
End Function

Public Function DoUrls(ByVal strText As String) As String
    Dim i As Integer
    Dim strLeft As String
    Dim strRight As String
    Dim strSub As String
    Dim strUrl As String
    Dim X As Integer
        
    'On Error GoTo exError
        
    strText = strText & " "
    strText = Replace(strText, "<", "<", , , vbTextCompare)
        
    strText = Replace(strText, "http://www.", "www.", , , vbTextCompare)
        
    i = InStr(1, strText, "www.", vbTextCompare)
    Do While Not i = 0
        strLeft = Left(strText, i - 1)
        If Not Right(strLeft, 9) = "<a href=" & Chr(34) And Not Right(strLeft, 2) = Chr(34) & ">" Then
            strRight = Mid(strText, i + (InStr(i, strText, " ") - i))
            strUrl = Mid(strText, i, (InStr(i, strText, " ") - i))
            strText = strLeft & "<a href=" & Chr(34) & "http://" & strUrl & Chr(34) & ">" & strUrl & "</a>" & strRight
            i = InStr((i + Len(strUrl)) + 20, strText, "www.", vbTextCompare)
        Else
            i = InStr(i + 1, strText, "www.", vbTextCompare)
        End If
    Loop
    
    i = InStr(1, strText, "http://", vbTextCompare)
    Do While Not i = 0
        strLeft = Left(strText, i - 1)
        If Not Right(strLeft, 9) = "<a href=" & Chr(34) And Not Right(strLeft, 2) = Chr(34) & ">" Then
            strRight = Mid(strText, i + (InStr(i, strText, " ") - i))
            strUrl = Mid(strText, i, (InStr(i, strText, " ") - i))
            strText = strLeft & "<a href=" & Chr(34) & strUrl & Chr(34) & ">" & strUrl & "</a>" & strRight
            i = InStr((i + Len(strUrl)) + 12, strText, "http://", vbTextCompare)
        Else
            i = InStr(i + 1, strText, "http://", vbTextCompare)
        End If
    Loop
        
exError:
    DoUrls = strText
End Function

Public Function ProcessSmileys(ByVal strText As String) As String
    Dim strList() As String
    Dim strSmiley() As String
    Dim strSmileys As String
    Dim strNum As String
    Dim intCount As Integer
    Dim X As Integer
        
    intCount = 10
    
    strSmileys = LoadTextFile(App.Path & "\Resources\Pictures\Smileys\Smileys.txt")
    strList = Split(strSmileys, vbCrLf)
    For X = 0 To UBound(strList)
        If Not Len(Trim(strList(X))) = 0 Then
            strSmiley = Split(strList(X), " ")
            strNum = Split(LCase(strSmiley(0)), "smiley")(1)
            If Not strSmiley(1) = "" Then
                If intCount <= 0 Then Exit For
                strSmiley(1) = Replace(strSmiley(1), "<", "&lt;")
                intCount = intCount - CountStr(strText, strSmiley(0))
                strText = Replace(strText, strSmiley(1), "<img src=" & Chr(34) & "file://" & App.Path & "\Resources\Pictures\Smileys\" & strNum & ".gif" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & ">", , 10, vbTextCompare)
            End If
            If UBound(strSmiley) > 1 Then
                If intCount <= 0 Then Exit For
                strSmiley(2) = Replace(strSmiley(2), "<", "&lt;")
                intCount = intCount - CountStr(strText, strSmiley(1))
                strText = Replace(strText, strSmiley(2), "<img src=" & Chr(34) & "file://" & App.Path & "\Resources\Pictures\Smileys\" & strNum & ".gif" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & ">", , 10, vbTextCompare)
            End If
        End If
    Next X
    ProcessSmileys = strText
End Function

Public Function CountStr(strString As String, strFind As String) As Integer
    Dim intFound As Integer
    Dim i As Integer
    i = InStr(strString, strFind)
    Do While Not i = 0
        intFound = intFound + 1
        i = InStr(i + 1, strString, strFind)
    Loop
    CountStr = intFound
End Function

Public Function RGB2Hex(ByVal strColor As ColorConstants) As String
    Dim Rb As Byte, Gb As Byte, Bb As Byte
    Dim R As String, G As String, B As String
    
    Bb = (strColor And 16711680) / 65536
    Gb = (strColor And 65280) / 256
    Rb = strColor And 255
    R = Hex(Rb)
    G = Hex(Gb)
    B = Hex(Bb)
    If Len(R) = 1 Then R = "0" & R
    If Len(G) = 1 Then G = "0" & G
    If Len(B) = 1 Then B = "0" & B
    
    RGB2Hex = "#" & R & G & B
End Function

Public Function Hex2RGB(ByVal strHex As String) As ColorConstants
    Dim R As String, G As String, B As String
    If Left(strHex, 1) = "#" Then strHex = Mid(strHex, 2)
    R = Mid(strHex, 1, 2)
    G = Mid(strHex, 3, 2)
    B = Mid(strHex, 5, 2)
    If Len(R) = 1 Then R = "0" & R
    If Len(G) = 1 Then G = "0" & G
    If Len(B) = 1 Then B = "0" & B
    R = Val("&h" & R)
    G = Val("&h" & G)
    B = Val("&h" & B)
    Hex2RGB = RGB(R, G, B)
End Function

Public Function Tag(strInside As String) As String
    Tag = Chr(34) & strInside & Chr(34)
End Function
