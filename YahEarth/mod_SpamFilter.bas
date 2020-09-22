Attribute VB_Name = "mod_SpamFilter"
Option Explicit
Type UsrLvl
    strUser As String
    strTime As Double
    intCount As Double
    blHigh As Boolean
End Type

Public tpUser() As UsrLvl
Public intUserCount As Double
Public intSpam As Double
Public intMessage As Double
Public strLastMsg As String

Public Function SpamFilter(strUser As String, strText As String) As Boolean
    Dim strFilter As String
    Dim strSpam() As String
    Dim X As Integer
    strFilter = LoadTextFile(App.Path & "\Resources\Filter\Spam.txt")
    
    If Len(strText) >= 5 And strText = strLastMsg Then GoTo exFound
        strLastMsg = strText
    
    If Options.blSpamFilter = False Then GoTo exOff
    If Not InStr(1, strText, strUser, vbTextCompare) = 0 Then
        strSpam = Split(strFilter, vbCrLf)
        For X = LBound(strSpam) To UBound(strSpam)
            If Not InStr(1, strText, strSpam(X), vbTextCompare) = 0 Then
                GoTo exFound
            End If
            DoEvents
        Next X
    Else
        If Options.blSpamAndUser = False Then
            strSpam = Split(strFilter, vbCrLf)
            For X = LBound(strSpam) To UBound(strSpam)
                If Not InStr(1, strText, strSpam(X), vbTextCompare) = 0 Then
                    GoTo exFound
                End If
                DoEvents
            Next X
        End If
    End If

exOff:
    SpamFilter = False
    Exit Function
exFound:
    SpamFilter = True
End Function

Public Function ParseWarning(strData As String)
    Dim strUser As String
    Dim strTime As Double
    Dim X As Integer
        
    If Not InStr(Replace(strData, "104À€", ""), "4À€") = 0 Then
        strUser = Parse("4À€", "À€", strData)
        If Len(strUser) < 3 Then GoTo exFound
        If InStr(strUser, ":") Then GoTo exFound
        'We can filter
        If intUserCount = 0 Then
            ReDim Preserve tpUser(intUserCount)
            intUserCount = intUserCount + 1
        End If
        
        For X = 0 To UBound(tpUser)
            If tpUser(X).strUser = strUser Then
                tpUser(X).intCount = tpUser(X).intCount + 1
                Debug.Print "User Warning (" & tpUser(X).strUser & "): " & tpUser(X).intCount
                strTime = Timer - tpUser(X).strTime
                If tpUser(X).intCount >= 3 Then
                    If strTime <= 1 Then 'User send more than 3 Packets per 1sec
                        'Ignore User
                        If tpUser(X).blHigh = False And IsIgnored(strUser) = False Then
                            AddIgnore strUser
                            ProcessError frm_Main.WB, "User '" & strUser & "' sent " & tpUser(X).intCount & " Packet's in " & strTime & "secs. User got automaticly Ignored", True
                            tpUser(X).blHigh = True
                        End If
                    Else
                        tpUser(X).intCount = 0
                        tpUser(X).strTime = Timer
                    End If
                End If
                GoTo exFound
            End If
        Next X

        ReDim Preserve tpUser(intUserCount)
        tpUser(intUserCount).strUser = strUser
        tpUser(intUserCount).strTime = Timer
        intUserCount = intUserCount + 1
    End If
exFound:
End Function

Public Function IsIgnored(strUser As String) As Boolean
    Dim strIgnores As String
    Dim strIgnored() As String
    Dim X As Integer
    
    strIgnores = LoadTextFile(App.Path & "\Resources\Filter\Ignored.txt")
    
    strIgnored = Split(strIgnores, vbCrLf)
    For X = 0 To UBound(strIgnored)
        If LTrim(RTrim(LCase(strUser))) = LTrim(RTrim(LCase(strIgnored(X)))) Then
            GoTo exFound
        End If
        DoEvents
    Next X
    
    IsIgnored = False
    Exit Function
exFound:
    IsIgnored = True
End Function

Public Function AddIgnore(strUser As String)
    Dim strIgnores As String
    strIgnores = LoadTextFile(App.Path & "\Resources\Filter\Ignored.txt")
    
    strIgnores = Replace(strIgnores, strUser & vbCrLf, "", , , vbTextCompare)
    
    If Right(strIgnores, 1) = vbCrLf Then
        strIgnores = strIgnores & strUser & vbCrLf
    Else
        If Len(Trim(strIgnores)) = 0 Then
            strIgnores = strUser & vbCrLf
        Else
            strIgnores = strIgnores & vbCrLf & strUser & vbCrLf
        End If
    End If
    
    SaveTextFile App.Path & "\Resources\Filter\Ignored.txt", strIgnores
End Function

Public Function RemoveIgnore(strUser As String)
    Dim strIgnores As String
    strIgnores = LoadTextFile(App.Path & "\Resources\Filter\Ignored.txt")
    
    strIgnores = Replace(strIgnores, strUser & vbCrLf, "", , , vbTextCompare)
        
    SaveTextFile App.Path & "\Resources\Filter\Ignored.txt", strIgnores
End Function
