Attribute VB_Name = "mod_Scripting"
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Type Sources
    strScript As String
    strFile As String
End Type
Public Src(8) As Sources
Public blScripting As Boolean
Public KeyWords As String

Public Function ExecuteScript(intCase As String, Optional strData As String, Optional strUser As String, Optional strMessage As String, Optional Window As String)
    Dim strHead As String
    Dim strFoot As String
    
    On Error Resume Next
    
    If Options.blScripting = False Then Exit Function
    
    Select Case intCase
        Case 1 'Incomming Data
            strHead = "Function IncommingData(Data)"
        Case 2 'Outgoing Data
            strHead = "Function OutgoingData(Data)"
        Case 3 'Incomming Chat Message
            strHead = "Function IncommingChatText(User, Message)"
        Case 4 'Incomming PM
            strHead = "Function IncommingPM(User, Message)"
        Case 5 'ApplicationStart
            strHead = "Function AppStart()"
        Case 6 'AppEnd
            strHead = "Function AppEnd()"
        Case 7 'New Window
            strHead = "Function NewWindow(Window)"
    End Select
    strFoot = "End Function"
    
    AddScriptObjects
    
    With frm_Main.Script
        .AddCode strHead & vbCrLf & Src(intCase).strScript & vbCrLf & strFoot & vbCrLf & _
            vbCrLf & Src(8).strScript
        Select Case intCase
            Case 1 'Incomming Data
                .Run "IncommingData", strData
            Case 2 'Outgoing Data
                .Run "OutgoingData", strData
            Case 3
                .Run "IncommingChatText", strUser, strMessage
            Case 4 'Incomming PM
                .Run "IncommingPM", strUser, strMessage
            Case 5 'AppStart
                .Run "AppStart"
            Case 6 'Incomming PM
                .Run "AppEnd"
            Case 7 'Incomming PM
                .Run "NewWindow", Window
        End Select
    End With
End Function

Private Function AddScriptObjects()
    Dim clsScripting As New cls_Scripts
    With frm_Main.Script
        'Add Objects
        .Reset
        .AddObject "frm_Login", frm_Login
        .AddObject "frm_Main", frm_Main
        .AddObject "frm_Rooms", frm_Rooms
        .AddObject "frm_Scripting", frm_Scripting
        .AddObject "frm_Smileys", frm_Smileys
        .AddObject "frm_Splash", frm_Splash
        .AddObject "Script", clsScripting
    End With
End Function

Public Function GetLineNum(RTB As RichTextBox) As Integer
    Dim I As Double
    Dim strLines() As String
    Dim strText As String
    
    I = RTB.SelStart
    strText = Left(RTB.Text, I)
    
    If Not InStr(strText, vbCrLf) = 0 Then
        strLines = Split(strText, vbCrLf)
        GetLineNum = UBound(strLines) + 1
    Else
        GetLineNum = 1
    End If
End Function

