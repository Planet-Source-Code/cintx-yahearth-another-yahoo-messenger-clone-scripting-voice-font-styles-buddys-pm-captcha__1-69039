Attribute VB_Name = "mod_DataOut"
Option Explicit
Public blSend As Boolean

Public Function SendData(strData As String) As Boolean
    If frm_Login.Socket.State = sckConnected Then
        Do While blSend = True
            DoEvents
        Loop
        ExecuteScript 2, strData
        frm_Login.Socket.SendData strData
        blSend = True
        SendData = True
    Else
        YMSG.strUser = "Not Connected"
        YMSG.strKey = ""
        YMSG.blJoined = False
        YMSG.strPass = ""
        YMSG.strRoom = ""
        YMSG.strRoomSpace = ""
        YMSG.strVoiceKey = ""
        SendData = False
    End If
End Function
