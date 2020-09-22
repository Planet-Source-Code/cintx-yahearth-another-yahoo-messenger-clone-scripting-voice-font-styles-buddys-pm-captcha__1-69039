Attribute VB_Name = "mod_Packets"
Option Explicit

Public YMSG_VER As Byte

Public Function Header(strID As String, strPacket As String) As String
    If YMSG.strKey = "" Then YMSG.strKey = String(4, 0)
    Header = "YMSG" & Chr(0) & Chr(YMSG_VER) & String(2, 0) & Chr(Fix(Len(strPacket) / 256)) & Chr(Len(strPacket) Mod 256) & _
    Chr(0) & Chr("&h" & strID) & String(4, 0) & YMSG.strKey & strPacket
    Debug.Print "[OUT]: " & Replace(Header, Chr(0), ".")
End Function

'---- Chat & PM

Public Function JoinChat(strRoom As String, strUser As String) As String
    JoinChat = Header("98", "1À€" & strUser & "À€104À€" & strRoom & "À€129À€1600326597À€62À€2À€")
End Function

Public Function PreJoin(strUser As String) As String
    PreJoin = Header("96", "109À€" & strUser & "À€1À€" & strUser & "À€6À€abcdeÀ€98À€usÀ€135À€ym8.1.0.421À€")
End Function

Public Function SendChat(strUser As String, strRoom As String, strMessage As String) As String
    SendChat = Header("A8", "1À€" & strUser & "À€104À€" & strRoom & "À€117À€" & strMessage & "À€124À€1À€")
End Function

Public Function Typing(strUser As String, strTo As String) As String
    Typing = Header("4B", "49À€TYPINGÀ€1À€" & strUser & "À€14À€ À€13À€1À€5À€" & strTo & "À€")
End Function

Public Function SendPM(strUser As String, strTo As String, strMsg As String, Optional MSN As Boolean = False) As String
    Dim strMSN As String
    If MSN = True Then strMSN = "À€241À€2"
    SendPM = Header("06", "1À€" & strUser & "À€5À€" & strTo & strMSN & "À€14À€" & strMsg & "À€97À€1À€63À€;0À€64À€0À€206À€0À€")
End Function

Public Function Ignore(strUser As String, strWho As String) As String
    Ignore = Header("85", "1À€" & strUser & "À€13À€1À€302À€319À€300À€319À€7À€" & strWho & "À€301À€319À€303À€319À€")
End Function

Public Function Leave(strUser As String) As String
    Leave = Header("A0", "1À€" & strUser & "À€")
End Function

Public Function AddContact(strUser As String, strFrom As String, strGroup As String, strMessage As String, strTo As String) As String
    AddContact = Header("83", "1À€" & strUser & "À€7À€" & strTo & "À€14À€" & strMessage & "À€65À€" & strGroup & "À€")
End Function

'----- Status Packets

Public Function Status_Busy() As String
    Status_Busy = Header("C6", "10À€2À€19À€À€97À€1À€")
End Function

Public Function Status_SteppedOut() As String
    Status_SteppedOut = Header("C6", "10À€9À€19À€À€97À€1À€47À€1À€")
End Function

Public Function Status_BrB() As String
    Status_BrB = Header("C6", "10À€1À€19À€À€97À€1À€")
End Function

Public Function Status_NotAtDesk() As String
    Status_NotAtDesk = Header("C6", "10À€4À€19À€À€97À€1À€")
End Function

Public Function Status_OnPhone() As String
    Status_OnPhone = Header("C6", "10À€6À€19À€À€97À€1À€")
End Function

Public Function Status_Custom(strStatus As String) As String
    Status_Custom = Header("C6", "10À€99À€19À€" & strStatus & "À€97À€1À€47À€0À€187À€0À€")
End Function

Public Function Status_Invisible() As String
    Status_Invisible = Header("C5", "13À€2À€")
End Function

Public Function Status_Invisible2(strUser As String) As String
    Status_Invisible2 = Header("BA", "1À€" & strUser & "À€31À€3À€13À€1À€")
End Function

Public Function Status_Available() As String
    Status_Available = Header("C6", "10À€0À€19À€À€97À€1À€")
End Function

Public Function Status_Online() As String
    Status_Online = Header("C5", "13À€1À€")
End Function
