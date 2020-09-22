Attribute VB_Name = "mod_Other"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long
  Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Public Function GetXCursorPos() As Long
   Dim pt As POINTAPI
   GetCursorPos pt
   GetXCursorPos = pt.X
End Function

Public Function GetYCursorPos() As Long
   Dim pt As POINTAPI
   GetCursorPos pt
   GetYCursorPos = pt.y
End Function

Public Function ColorMenu(Frm As Form, Color As Long)
    Dim mi As MENUINFO
    With mi
        .cbSize = Len(mi)
        .fMask = MIM_BACKGROUND
        .hbrBack = CreateSolidBrush(Color)
        SetMenuInfo GetMenu(Frm.hwnd), mi
        .fMask = MIM_BACKGROUND Or Color
        .hbrBack = CreateSolidBrush(vbCyan)
        SetMenuInfo GetSubMenu(GetMenu(Frm.hwnd), 0), mi
    End With
    DrawMenuBar Frm.hwnd
End Function
  
Public Function Parse(strL As String, strR As String, strData As String) As String
    Dim strSub As String
    Dim i As Integer
    i = InStr(strData, strL)
    If Not i = 0 Then
        strSub = Mid(strData, i + Len(strL))
    End If
    i = InStr(strSub, strR)
    If Not i = 0 Then
        Parse = Left(strSub, i - 1)
    Else
        Parse = strSub
    End If
End Function

Public Function OpenUrl(strUrl As Variant)
    ShellExecute 0, "open", strUrl, "", App.Path, 1
End Function

Public Function LoadTextFile(strFilename As String) As String
    Dim F As Integer
    Dim strContent As String
    Dim strBuffer As String
    
    On Error GoTo exError
    
    F = FreeFile
    
    Open strFilename For Input As #F
        Do While Not EOF(F)
            Line Input #F, strBuffer
            strContent = strContent & strBuffer & vbCrLf
            DoEvents
        Loop
    Close #F
    
    LoadTextFile = strContent
    
exError:
End Function

Public Function SaveTextFile(strFilename As String, strContent As String)
    Dim F As Integer
    
    On Error GoTo exError
    
    F = FreeFile
    
    Open strFilename For Output As #F
        Print #F, strContent
    Close #F
exError:
End Function

Public Function Status(intPanel As Integer, strText As String)
    frm_Main.StatusBar1.Panels(intPanel).Text = strText
End Function

Sub Pause(Seconds As Single)
    Dim Timer1 As Single, Timer2 As Single, currentDate As Date
    currentDate = Date
    Timer1 = Timer + Seconds
    Timer2 = Timer1 - 86400
    While ((Timer() < Timer1) And (currentDate = Date)) Or _
        ((Timer() < Timer2) And (currentDate + 1 = Date))
        DoEvents
    Wend
End Sub

Public Function FileCalc(Size As Double) As String
    On Error Resume Next
    Dim SizeArray(3) As String
    Dim Prefix As Integer
    Dim strSize As String
    Prefix = Int(Log(Size) / Log(1024))
    SizeArray(0) = "Bytes"
    SizeArray(1) = "KB"
    SizeArray(2) = "MB"
    SizeArray(3) = "GB"
    strSize = Str(Round(Size / (1024 ^ Prefix), 2)) & " " & SizeArray(Prefix)
    FileCalc = strSize
End Function

Public Function ValToBool(intVal As Integer) As Boolean
    If intVal = 1 Then ValToBool = True
    If intVal = 0 Then ValToBool = False
End Function

Public Function OpenFile(ListBox As ListBox, filename As String)
    Dim i As Integer, Str As String
    i = FreeFile
    If Not filename = "" Then
        Open filename For Input As #i
            Do While Not EOF(i)
                Line Input #i, Str
                If Not Str = "" Or Left(Str, 1) = "#" Or Str = " " Then ListBox.AddItem Str
                DoEvents
            Loop
        Close #i
    End If
End Function

Public Function SaveFile(ListBox As ListBox, filename As String)
    Dim i As Integer
    Dim X As Integer
    
    i = FreeFile
    If Not filename = "" Then
        Open filename For Output As #i
            For X = 0 To ListBox.ListCount - 1
                If Not ListBox.List(X) = "" Then Print #i, ListBox.List(X)
                DoEvents
            Next
        Close #i
    End If
End Function

Public Function DoCommand(strCommand As String)
    Dim strCase As String
    Dim strArg As String
    Dim i As Integer
    
    If Not Left(strCommand, 1) = "/" Then Exit Function
        
    i = InStr(strCommand, " ")
    If Not i = 0 Then
        strCase = Mid(Left(strCommand, i - 1), 2)
        strArg = Mid(strCommand, i + 1)
    Else
        Exit Function
    End If
    
    Select Case LCase(strCase)
        Case "join"
            DoReJoin strArg
    End Select
End Function

Function DoReJoin(strRoom As String)
    If YMSG.blJoined = True Then
        YMSG.strRoom = strRoom
        blRejoin = True
        SendData Leave(YMSG.strUser)
    Else
        YMSG.strRoom = strRoom
        SendData PreJoin(YMSG.strUser)
    End If
End Function

Public Function ConvertTimeStamp(dbStamp As Double) As String
    ConvertTimeStamp = TimestampToDate(dbStamp)
End Function
