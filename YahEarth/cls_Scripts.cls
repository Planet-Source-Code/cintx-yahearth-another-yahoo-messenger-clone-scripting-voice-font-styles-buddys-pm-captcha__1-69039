VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Scripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function SendChatMessage(ByVal strMessage As String)
    SendData mod_Packets.SendChat(YMSG.strUser, YMSG.strRoom, "<font INF ID:YahEarth Proto:YMSG VER:1.0PRV>" & strMessage)
    ProcessHTML YMSG.strUser, strMessage, frm_Main.WB
End Function

Function SendPM(ByVal strTo As String, ByVal strMessage As String)
    Dim X As Integer
    SendData mod_Packets.SendPM(YMSG.strUser, strTo, strMessage)
    X = FindPm(strTo, True)
    If Not X = 0 Then ProcessHTML YMSG.strUser, strMessage, PM(X).WB
End Function

Function Pause(ByVal Seconds As Single)
    Dim Timer1 As Single, Timer2 As Single, currentDate As Date
    currentDate = Date
    Timer1 = Timer + Seconds
    Timer2 = Timer1 - 86400
    While ((Timer() < Timer1) And (currentDate = Date)) Or _
        ((Timer() < Timer2) And (currentDate + 1 = Date))
        DoEvents
    Wend
End Function

Function Status(ByVal intPanel As Integer, ByVal strText As String)
    frm_Main.StatusBar1.Panels(intPanel).Text = strText
End Function

Function Parse(ByVal strL As String, ByVal strR As String, ByVal strData As String) As String
    Dim strSub As String
    Dim I As Integer
    I = InStr(strData, strL)
    strSub = Mid(strData, I + Len(strL))
    I = InStr(strSub, strR)
    Parse = Left(strSub, I - 1)
End Function

Function LoadTextFile(ByVal strFilename As String) As String
    Dim F As Integer
    Dim strContent As String
    Dim strBuffer As String
    
    On Error GoTo exError
    
    F = FreeFile
    
    Open strFilename For Input As #F
        Do While Not EOF(F)
            Line Input #F, strBuffer
            strContent = strContent & strBuffer & vbCrLf
            DoEvents
        Loop
    Close #F
    
    LoadTextFile = strContent
    
exError:
End Function

Function SaveTextFile(ByVal strFilename As String, ByVal strContent As String)
    Dim F As Integer
    
    On Error GoTo exError
    
    F = FreeFile
    
    Open strFilename For Output As #F
        Print #F, strContent
    Close #F
exError:
End Function

Function StripHtml(ByVal strText As String) As String
    Dim I As Integer
    Dim strLeft As String
    Dim strRight As String
    Dim strSub As String
    Dim X As Integer
    
    I = InStr(strText, "<")
    Do While Not I = 0
        strLeft = Left(strText, I - 1)
        strSub = Mid(strText, I)
        X = InStr(strSub, ">")
        If X = 0 Then
        
        End If
        strRight = Mid(strText, X + 1)
        strText = strLeft & " " & strRight
        
        I = InStr(I + 1, strText, "<")
        DoEvents
    Loop
    
    StripHtml = strText
End Function
