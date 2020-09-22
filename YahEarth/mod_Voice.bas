Attribute VB_Name = "mod_Voice"
Public blVoice As Boolean
Public Function DoVoice()
    Dim X As Integer
    
    On Error Resume Next
    
    If blVoice = False Then
        If YMSG.strVoiceKey = "" Then
            MsgBox "Error: Voice Key is empty", vbOKOnly, "Error"
            Exit Function
        End If
        If YMSG.strRoomSpace = "" Then
            MsgBox "Error: Roomspace is empty", vbOKOnly, "Error"
            Exit Function
        End If
        With frm_Main.Voice
            .leaveConference
            .HostName = "v4.vc.scd.yahoo.com"
            .loadSound YMSG.strUser
            .confKey = YMSG.strVoiceKey
            .confName = "ch/" & YMSG.strRoom & "::" & YMSG.strRoomSpace
            .Username = YMSG.strUser
            .appInfo = "mc(5, 6, 0, 1356)&u=" & YMSG.strUser & "&ia=us&lib=yacscom(45)&in=4,1,104,5.10"
            .createAndJoinConference
            .joinConference
        End With
    Else
        With frm_Main.Voice
            .leaveConference
        End With
        For X = 1 To frm_Main.lst_User.ListItems.Count
            frm_Main.lst_User.ListItems(X).SmallIcon = 1
        Next X
        blVoice = False
    End If
End Function

Public Function DoPMVoice(i As Integer)
    Dim X As Integer
    'With PM(i).Voice
        '.leaveConference
        '.HostName = "v4.vc.scd.yahoo.com"
        '.loadSound YMSG.strUser
        '.confKey = PMi(i).strVoiceKey
        '.confName = "pm/" & YMSG.strRoom & "::" & YMSG.strRoomSpace
        '.Username = YMSG.strUser
        '.appInfo = "mc(5, 6, 0, 1356)&u=" & YMSG.strUser & "&ia=us&lib=yacscom(45)&in=4,1,104,5.10"
        '.createAndJoinConference
        '.joinConference
    'End With
End Function
