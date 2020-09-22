Attribute VB_Name = "mod_PM"
Public Declare Function FlashWindow Lib "USER32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetForegroundWindow Lib "USER32" () As Long

Type PMInfo
    strTo As String
    blUsed As Boolean
    strVoiceKey As String
End Type
Public PMi(200) As PMInfo
Public PM(200) As New frm_PM
Public intPmCount As Integer

Public Function FindPm(strUser As String, Optional blTyping As Boolean = False, Optional blFocus As Boolean = False) As Integer
    Dim X As Integer
    
    For X = 1 To UBound(PMi)
        If PMi(X).strTo = strUser Then
            If PM(X).Visible = True Then
                If blFocus = True Then PM(X).SetFocus
                FindPm = X
                GoTo exFound
            Else
                PMi(X).blUsed = False
                PMi(X).strTo = ""
            End If
        End If
    Next X
    
    If blTyping = True Then
        FindPm = 0
        GoTo exFound
    Else
        For X = 1 To UBound(PMi)
            If PMi(X).blUsed = False Then
                If PM(X).Visible = False Then
                    InitPM X
                    PM(X).Visible = True
                    PMi(X).blUsed = True
                    PMi(X).strTo = strUser
                    PM(X).Tag = X
                    FindPm = X
                    intPmCount = intPmCount + 1
                    PM(X).Caption = strUser & " - " & YMSG.strUser
                    GoTo exFound
                Else
                    PMi(X).strTo = ""
                    PMi(X).blUsed = False
                End If
            End If
        Next X
    End If
exFound:
End Function

Public Function InitPM(intPM As Integer)
    PM(intPM).InitWindow
    PM(intPM).LoadWindow
    PM(intPM).LoadFonts
    PM(intPM).LoadImages
End Function
