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
    JoinChat = Header("98", "1��" & strUser & "��104��" & strRoom & "��129��1600326597��62��2��")
End Function

Public Function PreJoin(strUser As String) As String
    PreJoin = Header("96", "109��" & strUser & "��1��" & strUser & "��6��abcde��98��us��135��ym8.1.0.421��")
End Function

Public Function SendChat(strUser As String, strRoom As String, strMessage As String) As String
    SendChat = Header("A8", "1��" & strUser & "��104��" & strRoom & "��117��" & strMessage & "��124��1��")
End Function

Public Function Typing(strUser As String, strTo As String) As String
    Typing = Header("4B", "49��TYPING��1��" & strUser & "��14�� ��13��1��5��" & strTo & "��")
End Function

Public Function SendPM(strUser As String, strTo As String, strMsg As String, Optional MSN As Boolean = False) As String
    Dim strMSN As String
    If MSN = True Then strMSN = "��241��2"
    SendPM = Header("06", "1��" & strUser & "��5��" & strTo & strMSN & "��14��" & strMsg & "��97��1��63��;0��64��0��206��0��")
End Function

Public Function Ignore(strUser As String, strWho As String) As String
    Ignore = Header("85", "1��" & strUser & "��13��1��302��319��300��319��7��" & strWho & "��301��319��303��319��")
End Function

Public Function Leave(strUser As String) As String
    Leave = Header("A0", "1��" & strUser & "��")
End Function

Public Function AddContact(strUser As String, strFrom As String, strGroup As String, strMessage As String, strTo As String) As String
    AddContact = Header("83", "1��" & strUser & "��7��" & strTo & "��14��" & strMessage & "��65��" & strGroup & "��")
End Function

'----- Status Packets

Public Function Status_Busy() As String
    Status_Busy = Header("C6", "10��2��19����97��1��")
End Function

Public Function Status_SteppedOut() As String
    Status_SteppedOut = Header("C6", "10��9��19����97��1��47��1��")
End Function

Public Function Status_BrB() As String
    Status_BrB = Header("C6", "10��1��19����97��1��")
End Function

Public Function Status_NotAtDesk() As String
    Status_NotAtDesk = Header("C6", "10��4��19����97��1��")
End Function

Public Function Status_OnPhone() As String
    Status_OnPhone = Header("C6", "10��6��19����97��1��")
End Function

Public Function Status_Custom(strStatus As String) As String
    Status_Custom = Header("C6", "10��99��19��" & strStatus & "��97��1��47��0��187��0��")
End Function

Public Function Status_Invisible() As String
    Status_Invisible = Header("C5", "13��2��")
End Function

Public Function Status_Invisible2(strUser As String) As String
    Status_Invisible2 = Header("BA", "1��" & strUser & "��31��3��13��1��")
End Function

Public Function Status_Available() As String
    Status_Available = Header("C6", "10��0��19����97��1��")
End Function

Public Function Status_Online() As String
    Status_Online = Header("C5", "13��1��")
End Function
