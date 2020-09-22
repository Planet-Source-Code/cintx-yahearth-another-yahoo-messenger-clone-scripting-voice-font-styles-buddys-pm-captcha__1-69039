Attribute VB_Name = "mod_Login"
Option Explicit
Private Declare Function YCrypt Lib "YCrypt.dll" (ByVal Username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean
Private Type Y
    strUser As String
    strPass As String
    strKey As String
    strRoom As String
    blJoined As Boolean
    strRoomSpace As String
    strVoiceKey As String
End Type
Private Type C
    strUser As String
    strLastMsg As String
End Type
Public Chat() As C
Public YMSG As Y

Public Function GetHash(strUser As String) As String
    GetHash = Header(57, "1À€" & strUser & "À€")
End Function

Public Function Login(strUser As String, strPass As String, strHash As String) As String
    Dim strBuffer(1) As String, X As Integer

    'Fill the Buffer with Hex(00)
    strBuffer(0) = String(80, Chr(0))
    strBuffer(1) = String(80, Chr(0))
    
    If (YCrypt(strUser, strPass, strHash, strBuffer(0), strBuffer(1), 1) = False) Then
        'Failed to Get Encryption Keys, probably missing file
    Else
        'Encryption Keys are made able to Login now
        'Parse back the Buffer
        For X = 0 To 1
            strBuffer(X) = Left$(strBuffer(X), InStr(1, strBuffer(X), Chr(0)) - 1)
        Next X
        
        'Packet
        Login = "6À€" & strBuffer(0) & "À€96À€" & strBuffer(1) & "À€0À€" & strUser & "À€2À€1À€1À€" & strUser & _
        "À€135À€5, 6, 0, 1347À€148À€300À€" 'Send Login Packet with Keys,Version (Major, Minor, Revision, Build)
        Login = Header(54, Login)
    End If
End Function

Public Function ClearFields()
    With YMSG
        .blJoined = False
        .strKey = ""
        .strPass = ""
        .strRoom = ""
        .strRoomSpace = ""
        .strUser = ""
        .strVoiceKey = ""
    End With
End Function
