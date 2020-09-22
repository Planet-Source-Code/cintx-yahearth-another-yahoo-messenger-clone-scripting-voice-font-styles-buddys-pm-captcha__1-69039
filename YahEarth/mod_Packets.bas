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
    JoinChat = Header("98", "1¿Ä" & strUser & "¿Ä104¿Ä" & strRoom & "¿Ä129¿Ä1600326597¿Ä62¿Ä2¿Ä")
End Function

Public Function PreJoin(strUser As String) As String
    PreJoin = Header("96", "109¿Ä" & strUser & "¿Ä1¿Ä" & strUser & "¿Ä6¿Äabcde¿Ä98¿Äus¿Ä135¿Äym8.1.0.421¿Ä")
End Function

Public Function SendChat(strUser As String, strRoom As String, strMessage As String) As String
    SendChat = Header("A8", "1¿Ä" & strUser & "¿Ä104¿Ä" & strRoom & "¿Ä117¿Ä" & strMessage & "¿Ä124¿Ä1¿Ä")
End Function

Public Function Typing(strUser As String, strTo As String) As String
    Typing = Header("4B", "49¿ÄTYPING¿Ä1¿Ä" & strUser & "¿Ä14¿Ä ¿Ä13¿Ä1¿Ä5¿Ä" & strTo & "¿Ä")
End Function

Public Function SendPM(strUser As String, strTo As String, strMsg As String, Optional MSN As Boolean = False) As String
    Dim strMSN As String
    If MSN = True Then strMSN = "¿Ä241¿Ä2"
    SendPM = Header("06", "1¿Ä" & strUser & "¿Ä5¿Ä" & strTo & strMSN & "¿Ä14¿Ä" & strMsg & "¿Ä97¿Ä1¿Ä63¿Ä;0¿Ä64¿Ä0¿Ä206¿Ä0¿Ä")
End Function

Public Function Ignore(strUser As String, strWho As String) As String
    Ignore = Header("85", "1¿Ä" & strUser & "¿Ä13¿Ä1¿Ä302¿Ä319¿Ä300¿Ä319¿Ä7¿Ä" & strWho & "¿Ä301¿Ä319¿Ä303¿Ä319¿Ä")
End Function

Public Function Leave(strUser As String) As String
    Leave = Header("A0", "1¿Ä" & strUser & "¿Ä")
End Function

Public Function AddContact(strUser As String, strFrom As String, strGroup As String, strMessage As String, strTo As String) As String
    AddContact = Header("83", "1¿Ä" & strUser & "¿Ä7¿Ä" & strTo & "¿Ä14¿Ä" & strMessage & "¿Ä65¿Ä" & strGroup & "¿Ä")
End Function

'----- Status Packets

Public Function Status_Busy() As String
    Status_Busy = Header("C6", "10¿Ä2¿Ä19¿Ä¿Ä97¿Ä1¿Ä")
End Function

Public Function Status_SteppedOut() As String
    Status_SteppedOut = Header("C6", "10¿Ä9¿Ä19¿Ä¿Ä97¿Ä1¿Ä47¿Ä1¿Ä")
End Function

Public Function Status_BrB() As String
    Status_BrB = Header("C6", "10¿Ä1¿Ä19¿Ä¿Ä97¿Ä1¿Ä")
End Function

Public Function Status_NotAtDesk() As String
    Status_NotAtDesk = Header("C6", "10¿Ä4¿Ä19¿Ä¿Ä97¿Ä1¿Ä")
End Function

Public Function Status_OnPhone() As String
    Status_OnPhone = Header("C6", "10¿Ä6¿Ä19¿Ä¿Ä97¿Ä1¿Ä")
End Function

Public Function Status_Custom(strStatus As String) As String
    Status_Custom = Header("C6", "10¿Ä99¿Ä19¿Ä" & strStatus & "¿Ä97¿Ä1¿Ä47¿Ä0¿Ä187¿Ä0¿Ä")
End Function

Public Function Status_Invisible() As String
    Status_Invisible = Header("C5", "13¿Ä2¿Ä")
End Function

Public Function Status_Invisible2(strUser As String) As String
    Status_Invisible2 = Header("BA", "1¿Ä" & strUser & "¿Ä31¿Ä3¿Ä13¿Ä1¿Ä")
End Function

Public Function Status_Available() As String
    Status_Available = Header("C6", "10¿Ä0¿Ä19¿Ä¿Ä97¿Ä1¿Ä")
End Function

Public Function Status_Online() As String
    Status_Online = Header("C5", "13¿Ä1¿Ä")
End Function
