Attribute VB_Name = "mod_Buddys"
Public strGroup As String
Public strLastUser As String
Public blLastUserState As Boolean

Public Function DoBuddy(strUser As String, lst_Buddy As TreeView, Optional blOnline As Boolean = False, Optional blAdd As Boolean = True) As String
    Dim X As Integer
    Dim strRoot As String
        
    If lst_Buddy.Nodes.Count = 0 Then
        lst_Buddy.Nodes.Add , , "Online", "Online", 4
        lst_Buddy.Nodes.Add , , "Offline", "Offline", 3
        lst_Buddy.Nodes(1).Expanded = True
        'lst_Buddy.Nodes(2).Expanded = True
    End If
        
    For X = 1 To lst_Buddy.Nodes.Count
        If Mid(LCase(lst_Buddy.Nodes(X).Key), 3) = LCase(strUser) Then
            lst_Buddy.Nodes.Remove X
            If blOnline = True Then
                lst_Buddy.Nodes.Add "Online", tvwChild, "u_" & strUser, strUser, 1
            Else
                lst_Buddy.Nodes.Add "Offline", tvwChild, "u_" & strUser, strUser, 2
            End If
            DoBuddy = strUser
            Exit Function
        End If
    Next X
    
    If blAdd = True Then
        If blOnline = True Then
            lst_Buddy.Nodes.Add "Online", tvwChild, "u_" & strUser, strUser, 1
        Else
            lst_Buddy.Nodes.Add "Offline", tvwChild, "u_" & strUser, strUser, 2
        End If
        DoBuddy = strUser
    End If
    
    SetCount lst_Buddy
End Function

Public Function SetStatus(strUser As String, lst_Buddy As TreeView, Optional strStatus As String = "", Optional blOnline As Boolean = True)
    Dim X As Integer
    
    For X = 1 To lst_Buddy.Nodes.Count
        If Mid(LCase(lst_Buddy.Nodes(X).Key), 3) = LCase(strUser) Then
            If blOnline = False Then DoBuddy strUser, lst_Buddy, False, False
            If strStatus = "" Then
                lst_Buddy.Nodes(X).Text = strUser
            Else
                lst_Buddy.Nodes(X).Text = strUser & " - " & strStatus
            End If
        End If
    Next X
End Function

Public Function SetStatus2(strUser As String, lst_Buddy As ListView, blOnline As Boolean)
    Dim X As Integer
    
    For X = 1 To lst_Buddy.ListItems.Count
        If LCase(strUser) = LCase(lst_Buddy.ListItems(X).Text) Then
            If blOnline = False Then
                lst_Buddy.ListItems(X).SmallIcon = 2
            Else
                lst_Buddy.ListItems(X).SmallIcon = 1
            End If
            Exit Function
        End If
    Next X
End Function

Public Function SetCount(lst_Buddy As TreeView)
    Dim intOnline As Integer
    Dim intOffline As Integer
    Dim X As Integer
    
    For X = 1 To lst_Buddy.Nodes.Count
        If Left(lst_Buddy.Nodes(X).FullPath, Len("Online")) = "Online" Then
            intOnline = intOnline + 1
        End If
                
        If Left(lst_Buddy.Nodes(X).FullPath, Len("Offline")) = "Offline" Then
            intOffline = intOffline + 1
        End If
    Next X
    
    lst_Buddy.Nodes(1).Text = "Online (" & (intOnline - 1) & ")"
    lst_Buddy.Nodes(2).Text = "Offline (" & (intOffline - 1) & ")"
End Function
