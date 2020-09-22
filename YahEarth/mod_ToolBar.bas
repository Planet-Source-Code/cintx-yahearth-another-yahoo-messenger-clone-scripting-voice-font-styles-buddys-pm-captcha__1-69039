Attribute VB_Name = "mod_ToolBar"
Option Explicit
Public Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Const VK_CONTROL = &H11
Public Const VK_C = &H43
Public Const KEYEVENTF_KEYUP = &H2

Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
    
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_RBUTTONDOWN = &H204

Public origWndProc As Long

'Toolbar
Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Function AppWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_MOUSEACTIVATE
            Dim C As Integer
            Call CopyMemory(C, ByVal VarPtr(lParam) + 2, 2)
            If C = WM_RBUTTONDOWN Then
                'frm_PM.PopupMenu frm_Main.mnu_Edit
                SendKeys "{ESC}"
            End If
    End Select
    AppWndProc = CallWindowProc(origWndProc, hWnd, Msg, wParam, lParam)
End Function

Sub ToolFlat(ControlName As Control, flat As Boolean)
    Dim style As Long
    Dim hToolbar As Long
    Dim r As Long
       
'Now Make it Flat
    'First get the hWnd
    hToolbar = FindWindowEx(ControlName.hWnd, 0&, "ToolbarWindow32", vbNullString)
    'get Style
    style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
    'Change style
    If (style And TBSTYLE_FLAT) And Not flat Then
        style = style Xor TBSTYLE_FLAT
    ElseIf flat Then
        style = style Or TBSTYLE_FLAT
    End If
    'Set the Style
    r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, style)
    'Now show what we've done, this isn't neccesary if used in form_load
    ControlName.Refresh
End Sub

