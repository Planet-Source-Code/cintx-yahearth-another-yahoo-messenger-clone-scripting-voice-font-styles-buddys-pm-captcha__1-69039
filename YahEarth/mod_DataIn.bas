Attribute VB_Name = "mod_DataIn"
Option Explicit
Public strCount As String
Public blRejoin As Boolean
Public strDataBuffer As String
Public intSizeBuffer As Long

Public Function IncommingData(ByVal strData As String) As Boolean
    'Just for Data \/
    
    On Error Resume Next
    
    Dim i As Integer
    Dim intPacketLength As Long
    Dim strNextData As String
        
    If Not Mid(strData, 1, 4) = "YMSG" Then
        'Data does not start with YMSG
        If Not strDataBuffer = "" Then
            strDataBuffer = strDataBuffer & strData
            If intSizeBuffer <= Len(strDataBuffer) Then
                strData = strDataBuffer
                If intSizeBuffer < Len(strDataBuffer) Then
                    strNextData = Mid(strDataBuffer, intSizeBuffer + 1)
                    strData = left(strDataBuffer, intSizeBuffer)
                    PacketIdentifier strData
                    IncommingData strNextData
                ElseIf intSizeBuffer = Len(strDataBuffer) Then
                    intSizeBuffer = 0
                    strDataBuffer = ""
                    PacketIdentifier strData
                ElseIf intSizeBuffer > Len(strDataBuffer) Then
                    strDataBuffer = strDataBuffer & strData
                End If
            ElseIf Len(strDataBuffer) < intSizeBuffer Then
                DoEvents
                Exit Function
            End If
        End If
    Else
        'Data starts with YMSG
        intPacketLength = Val(Asc(Mid(strData, 9, 1))) * 256 + Val(Asc(Mid(strData, 10, 1))) + 20 'Packet Length
        If intPacketLength > Len(strData) Then
            'Missing Data wait till new Data comes
            strDataBuffer = strData
            intSizeBuffer = intPacketLength
            Exit Function
        ElseIf intPacketLength <= Len(strData) Then
            If intPacketLength < Len(strData) Then 'More than 1 Packet, recall
                strNextData = Mid(strData, intPacketLength + 1)
                strData = left(strData, intPacketLength)
                PacketIdentifier strData
                IncommingData strNextData
            ElseIf intPacketLength = Len(strData) Then
                strData = left(strData, intPacketLength)
                PacketIdentifier strData
                Exit Function
            End If
        End If
    End If
End Function
Public Function PacketIdentifier(strData As String)
    Dim strID As Byte
    
    If strData = "" Then Exit Function
    
    strID = Asc(Mid(strData, 12, 1)) 'Protocol Type
    
    Debug.Print "[" & strID & String(3 - Len(strID), " ") & "]: " & Replace(strData, Chr(0), ".")
    
    Select Case strID 'Packet Case
        Case 1
            'YMSG12 Buddy Status
            ParseStatus strData, frm_Buddys.lst_Buddy, True
        Case 2
            'YMSG12 Buddy Status
            ParseStatus strData, frm_Buddys.lst_Buddy, False
        Case 6
            'Received PM
            ParsePM strData, frm_Main.WB
        Case 11
            'Dunno
            
        Case 15
            'New Buddy
            NewBuddy strData
        Case 18
            'Dunno
            
        Case 75
            'User is Typing
            ParseTyping strData
        Case 79
            'Peer to Peer (eh)?
            
        Case 87
            'Send Login
            SendLogin strData
        Case 84
            'Invalid Password
            InvalidLogin strData
        Case 85
            'Logged In, Buddys
            LoggedIn strData
            ParseBuddys strData, frm_Buddys.lst_Buddy
        Case 150
            'Ready to Join Chat
            JoinRoom YMSG.strRoom, strData
        Case 152
            'Joined Chat
            ParseRoom strData, frm_Main.lst_User, frm_Main.WB
        Case 155
            'Left Room
            DeParseRoom strData, frm_Main.lst_User, frm_Main.WB
        Case 160
            ParseLeaveChat strData
        Case 168
            'Chat Text arrival
            ParseChat strData, frm_Main.WB
        Case 186
            SendData Status_Invisible
        Case 198
            'New User Status
            ParseUserStatus strData, frm_Buddys.lst_Buddy
        Case 225
            'Yahoo 360 fuck this
            
        Case 239
            'still dunno
            
        Case 240
            'YMSG15 Buddy Status
            ParseStatus strData, frm_Buddys.lst_Buddy, True
            
        Case 241
            'YMSG15 Buddy List
            ParseBuddys_15 strData, frm_Buddys.lst_Buddy
        Case Else
            'Unknown Protocol
            ProcessError frm_Main.WB, "Unknown Protocol (Case: " & strID & ") Received"
    End Select
    
    ParseWarning strData
    
    If blScripting = True Then
        ExecuteScript 1, strData
    End If
End Function

Public Function InvalidLogin(strData As String)
    frm_Login.Socket.Close
    frm_Login.StatusBar1.Panels(1).Text = "Status: Invalid Password"
    Status 2, "Online Status: Offline"
End Function

Public Function SendLogin(strData As String)
    Dim strHash As String
    strHash = Parse("94��", "��", strData)
    If Not strHash = "" Then
        YMSG.strKey = Mid(strData, 17, 4)
        SendData Login(YMSG.strUser, YMSG.strPass, strHash)
    End If
End Function

Public Function LoggedIn(strData As String)
    Unload frm_Login
    frm_Main.mnu_Logout.Enabled = True
    frm_Main.mnu_Login.Enabled = False
    Status 2, "Online Status: Online"
    If frm_Login.check_Join.Value = 1 Then SendData PreJoin(YMSG.strUser)
End Function

Public Function JoinRoom(strRoom As String, strData As String)
    SendData JoinChat(YMSG.strRoom, YMSG.strUser)
End Function

Public Function DeParseRoom(strData As String, lst_User As ListView, WB As WebBrowser)
    Dim strUser As String

    strUser = Parse("109��", "��", strData)
    LeftRoom strUser, lst_User, WB
End Function

Public Function LeftRoom(strUser As String, lst_User As ListView, WB As WebBrowser)
    Dim X As Integer
    For X = 1 To lst_User.ListItems.Count
        If LCase(lst_User.ListItems(X).Text) = LCase(strUser) Then
            lst_User.ListItems.Remove X
            If YMSG.blJoined = True Then
                ProcessRoomUser strUser, WB, "left the Room"
            End If
            Exit For
        End If
    Next X
End Function

Public Function ParseRoom(strData As String, lst_User As ListView, WB As WebBrowser)
    Dim strCase As String
    
    If Not InStr(1, strData, "����", vbTextCompare) = 0 Then
        strCase = Parse("114��", "��", strData)
        If left(strCase, 2) = "-6" Then
            ProcessError WB, "Room not Found"
        ElseIf left(strCase, 3) = "-35" Then
            ProcessError WB, "Room is Full"
        ElseIf left(strCase, 2) = "16" Then
            'Never seen this
        Else
            ProcessError WB, "Unknown Chat Error"
        End If
    Else
        ParseChatList strData, lst_User, WB
    End If
End Function

Public Function ParseChatList(strData As String, lst_User As ListView, WB As WebBrowser)
    Dim strUsers() As String
    Dim strUser As String
    Dim strRoom As String
    Dim strTopic As String
    Dim strCaptcha As String
    Dim X As Integer
    
    'On Error Resume Next
    
    If YMSG.blJoined = False Then
        If strCount = "" Then
            strTopic = Parse("105��", "��", strData)
            
            ' Our Captcha !!11111oneoen
            ' I just updated the packets and see there it finally works again
            
            If Not InStr(1, strTopic, "captcha.chat.yahoo.com", vbTextCompare) = 0 Then
                strCaptcha = Mid(strTopic, InStr(1, strTopic, "http://", vbTextCompare))
                If Not strCaptcha = "" Then
                    Debug.Print strCaptcha
                    frm_Captcha.ShowCaptcha (strCaptcha)
                    Exit Function
                End If
            End If
            
            strCount = Parse("108��", "��", strData)
            strRoom = Parse("104��", "��", strData)
            frm_Main.Caption = "YahEarth - " & YMSG.strRoom
            ProcessRoom strRoom, strTopic, WB, strCount
            If YMSG.strRoomSpace = "" Then YMSG.strRoomSpace = Parse("129��", "��", strData)
            If YMSG.strVoiceKey = "" Then YMSG.strVoiceKey = Parse("130��", "��", strData)
        End If
    End If
    
    strUsers = Split(strData, "109��")
    For X = 1 To UBound(strUsers)
        strUser = Split(strUsers(X), "��")(0)
        If Not strUser = "" Then
            LeftRoom strUser, lst_User, WB
            lst_User.ListItems.Add , , strUser, 1, 1
            If YMSG.blJoined = True Then
                'Do Message for User Joined
                ProcessRoomUser strUser, WB, "joined the Room"
            End If
        End If
        DoEvents
    Next X
    
    If YMSG.blJoined = False Then
        If strCount <= lst_User.ListItems.Count Then
            YMSG.blJoined = True
            strCount = ""
            If frm_Login.check_Voice.Value = 1 Then DoVoice
            Status 1, "Status: Joined Room"
        End If
    End If
End Function

Public Function ParseChat(strData As String, WB As WebBrowser)
    Dim strUser As String
    Dim strMessage As String
    
    strUser = Parse("109��", "��", strData)
    strMessage = Parse("117��", "��", strData)
    ExecuteScript 3, , strUser, strMessage
    If Not SpamFilter(strUser, strMessage) = True Then
        If Not IsIgnored(strUser) = True Then
            ProcessHTML strUser, strMessage, WB
            intMessage = intMessage + 1
            Status 4, "Messages: " & intMessage
        End If
    Else
        intSpam = intSpam + 1
        Status 3, "Blocked Spam: " & intSpam
    End If
End Function

Public Function ParseUserOffline(strData As String, WB As WebBrowser)
    Dim strUser As String
    
    strUser = Parse("7��", "��", strData)
End Function

Public Function ParsePM(strData As String, WB As WebBrowser)
    Dim strUser As String
    Dim strMsg As String
    Dim X As Integer
    
    '[6  ]: YMSG.....}....�R�{5��yahearth_test��4��dear_matt_hew��206��2��252��yrGX8ewaqtVWdbkg6kkB7hhnxG5p5A==��97��1��14��HOMO YOU HOMO!��63��;0��64��0��
    If Not InStr(strData, "32��") = 0 Then
        ParseOffline strData
        Exit Function
    End If
    
    strUser = Parse("4��", "��", strData)
    strMsg = Parse("14��", "��", strData)
    
    If Not SpamFilter(strUser, strMsg) = True Then
        If Not IsIgnored(strUser) = True Then
            X = FindPm(strUser)
            If Not GetForegroundWindow = PM(X).hWnd Then
                'Play Sound
                'PlaySound App.Path & "\Resources\Sounds\message.mp3"
            End If
            
            If Not InStr(strData, "252��") = 0 Then
                If Not PMi(X).strVoiceKey = "" Then
                    PMi(X).strVoiceKey = Parse("252��", "��", strData)
                    PM(X).tlb_Buttons.Buttons(2).Enabled = True
                Else
                    If Not PMi(X).strVoiceKey = Parse("252��", "��", strData) Then
                        PMi(X).strVoiceKey = Parse("252��", "��", strData)
                    End If
                End If
            Else
                If Not PMi(X).strVoiceKey = "" Then
                    PM(X).tlb_Buttons.Buttons(2).Enabled = False
                End If
            End If
            FlashWindow PM(X).hWnd, 3
            ProcessHTML strUser, strMsg, PM(X).WB, True
            ExecuteScript 4, , strUser, strMsg
            PM(X).StatusBar1.Panels(1).Text = "Last message received on " & Time
        End If
    Else
        intSpam = intSpam + 1
        Status 3, "Blocked Spam: " & intSpam
    End If
End Function

Public Function ParseTyping(strData As String)
    Dim strUser As String
    Dim X As Integer
    
    '[75 ]: YMSG.....?.K...�U�5��yahearth_test��4��dear_matt_hew��14�� ��13��1��49��TYPING��
    
    strUser = Parse("4��", "��", strData)
    
    X = FindPm(strUser, True)
    If X = 0 Then Exit Function
    PM(X).StatusBar1.Panels(1).Text = strUser & " is typing a Message"
End Function

Public Function ParseLeaveChat(strData As String)
    If blRejoin = True Then
        YMSG.blJoined = False
        YMSG.strRoomSpace = ""
        YMSG.strVoiceKey = ""
        frm_Main.lst_User.ListItems.Clear
        frm_Main.InitWindow
        If blVoice = True Then DoVoice
        
        'Join Chat
        SendData PreJoin(YMSG.strUser)
        blRejoin = False
    End If
End Function

Public Function ParseBuddys(strData As String, lst_Buddy As TreeView)
    Dim strUsers() As String
    Dim strList As String
    Dim X As Integer
    
    '[85  ]: YMSG....�.U....�S�,87��Friends:dear_matt_hew,dosed
    '��88����89��yahearth_test��59��Y    v=1&n=4a5bjap2fqq7v&l=o0740hj7_j4ij/o&p=m2g0c58012000000&r=gj&lg=us&intl=us; expires=Thu, 15 Apr 2010 20:00:00 GMT; path=/; domain=.yahoo.com��219����59��T z=mLBGFBmRWGFBJeXKNp5FpR7Mk80BjYxNzA2Mzc0NE8-&a=QAE&sk=DAA1XaCpw/EJIF&d=c2wBTlRnekFURTJNRGN4TkRBek16Zy0BYQFRQUUBdGlwAVVkS1dERAF6egFtTEJHRkJnV0E-; expires=Thu, 15 Apr 2010 20:00:00 GMT; path=/; domain=.yahoo.com��219����59��C    mg=1��219����153��1��90��1��3��yahearth_test��100��0��101����102����15001��0��15002��us��213��0��275��1��216��matthew��254��yahearth��93��86400��149��q7owxjeFaNSmDwSyqZ54kA--��150��cboi9hxyiecXnwhpKFNqPA--��151��Fdc4liMyO.mJrhMRB5qvWQ--��217��0��.
    
    strList = Parse("87��", "��", strData)
    strList = Replace(Replace(strList, ":", ":,"), Chr(10), "")
    
    strUsers = Split(strList, ",")
    
    For X = 0 To UBound(strUsers)
        If Not Right(strUsers(X), 1) = ":" Then
            DoBuddy strUsers(X), lst_Buddy, False, True
            frm_NewPM.lst_Buddy.ListItems.Add , , strUsers(X), 2, 2
        End If
    Next X
End Function

Public Function ParseStatus(strData As String, lst_Buddy As TreeView, blOnline As Boolean)
    Dim strUsers() As String
    Dim strUser As String
    Dim strList As String
    Dim strStatus As String
    Dim X As Integer
    
    strUsers = Split(strData, "7��")
    For X = 1 To UBound(strUsers)
        strUser = Split(strUsers(X), "��")(0)
        strStatus = DoBuddy(strUser, lst_Buddy, blOnline, False)
        SetStatus2 strUser, frm_NewPM.lst_Buddy, blOnline
        If Not strLastUser = strUser Then
            If Not blOnline = blLastUserState Then
                strLastUser = strUser
                blLastUserState = blOnline
            End If
        Else
            If Not blOnline = blLastUserState Then
                strLastUser = strUser
                blLastUserState = blOnline
            End If
        End If
        ParseUserStatus "7��" & strUsers(X), lst_Buddy
    Next X
End Function

Public Function ParseUserStatus(strData As String, lst_Buddy As TreeView)
    Dim strUser As String
    Dim strStatus As String
    Dim strCustom As String
        
    strUser = Parse("7��", "��", strData)
    strStatus = Parse("10��", "��", strData)
    
    Select Case strStatus
        Case -1
            SetStatus strUser, lst_Buddy, , False
            SetStatus2 strUser, frm_NewPM.lst_Buddy, False
        Case 0
            SetStatus strUser, lst_Buddy, , True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
        Case 2
            SetStatus strUser, lst_Buddy, "Busy", True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
        Case 9
            SetStatus strUser, lst_Buddy, "Stepped Out", True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
        Case 1
            SetStatus strUser, lst_Buddy, "Be Right Back", True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
        Case 4
            SetStatus strUser, lst_Buddy, "Not at My Desk", True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
        Case 6
            SetStatus strUser, lst_Buddy, "On the Phone", True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
        Case 99
            strCustom = Parse("19��", "��", strData)
            SetStatus strUser, lst_Buddy, strCustom, True
            SetStatus2 strUser, frm_NewPM.lst_Buddy, True
    End Select
End Function

Public Function ParseBuddys_15(strData As String, lst_Buddy As TreeView)
    Dim strBuddy() As String
    Dim strUser As String
    Dim X As Integer
    
    strBuddy = Split(strData, "319��7��")
    
    For X = 1 To UBound(strBuddy)
        strUser = Split(strBuddy(X), "��")(0)
        If Not Len(strUser) < 2 Then
            DoBuddy strUser, lst_Buddy, False, True
            frm_NewPM.lst_Buddy.ListItems.Add , , strUser, 2, 2
        End If
    Next X
End Function

Public Function ParseOffline(strData As String)
    Dim strUser As String
    Dim strMsg As String
    Dim strTime As String
    Dim blOption As Boolean
    Dim X As Integer
    Dim strList() As String
    
    '[6  ]: YMSG....u....�@�31��6��32��6��5��dear_matt_hew��4��yahearth_test��15��1160239944��252��gcmkZ0J6YgFeXUeTYCVnJyuJKuMJoQ==��14��dfhdfhd��97��1��31��6��32��6��5��dear_matt_hew��4��yahearth_test��15��1160239945��252��FQXZ8o7N9MGfAIQKQD2JPlaSTxuJ7Q==��14��hdfhdh��97��1��31��6��32��6��5��dear_matt_hew��4��yahearth_test��15��1160239946��252��+rx/j4FgWlXbylmTcywYDDiC1c5POA==��14��dhdfhd��97��1��
    
    strList = Split(strData, "��32")
    blOption = Options.blDisableFontStyle
    Options.blDisableFontStyle = True
    For X = 1 To UBound(strList)
        strUser = Parse("4��", "��", strList(X))
        strMsg = Parse("14��", "��", strList(X))
        strTime = ConvertTimeStamp(Val(Parse("15��", "��", strList(X))))
        
        With frm_Offline.lst_Offline.ListItems.Add(, , strUser)
            .SubItems(1) = strTime
            .SubItems(2) = ProcessHTML(strUser, strMsg, frm_Offline.WB, True, True)
        End With
    Next X
    Options.blDisableFontStyle = blOption
    
    frm_Offline.Show
End Function

Public Function NewBuddy(strData As String)
    Dim strUser As String
    strUser = Parse("7��", "��", strData)
    DoBuddy strUser, frm_Buddys.lst_Buddy, False, True
    frm_NewPM.lst_Buddy.ListItems.Add , , strUser, 2, 2
    ParseUserStatus strData, frm_Buddys.lst_Buddy
End Function
