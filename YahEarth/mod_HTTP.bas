Attribute VB_Name = "mod_HTTP"
Public Function ParseUrl(ByVal strUrl As String, _
                          ByRef strHost As String, _
                          ByRef strPort As String, _
                          ByRef strRequest As String)
    Dim I As Integer
    Dim X As Integer
    Dim strSub As String
    
    If LCase(Left(strUrl, Len("http://"))) = "http://" Then
        'http:// exists?
        strUrl = Mid(strUrl, Len("http://") + 1)
    End If
    
    If LCase(Left(strUrl, Len("www."))) = "www." Then
        'www. exists?
        strUrl = Mid(strUrl, Len("www.") + 1)
    End If
    
    I = InStr(strUrl, "/")
    If Not I = 0 Then
        strRequest = Mid(strUrl, I)
        strSub = Left(strUrl, I - 1)
        I = InStr(strSub, ":")
        If Not I = 0 Then
            strPort = Mid(strSub, I + 1)
            strHost = Left(strSub, I - 1)
        Else
            strPort = 80
            strHost = strSub
        End If
    Else
        strRequest = "/"
        I = InStr(strUrl, ":")
        If Not I = 0 Then
            strPort = Mid(strUrl, I + 1)
            strHost = Left(strUrl, I - 1)
        Else
            strPort = 80
            strHost = strUrl
        End If
    End If
End Function

