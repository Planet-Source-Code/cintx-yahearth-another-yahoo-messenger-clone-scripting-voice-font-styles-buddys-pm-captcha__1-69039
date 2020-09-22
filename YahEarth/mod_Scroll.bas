Attribute VB_Name = "mod_Scroll"
Private Declare Function GetScrollBarInfo Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal idObject As Long, _
                psbi As SCROLLBARINFO) As Long
                
Private Declare Function GetScrollInfo Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal n As Long, _
                lpScrollInfo As SCROLLINFO) As Long
                
Private Declare Function SetScrollInfo Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal n As Long, _
                lpcScrollInfo As SCROLLINFO, _
                ByVal bool As Boolean) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
                
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Type SCROLLBARINFO
    cbSize As Long
    rcScrollBar As RECT
    dxyLineButton As Long
    xyThumbTop As Long
    xyThumbBottom As Long
    reserved As Long
    rgstate(0 To 5) As Long
End Type

Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS)

Private Const OBJID_VSCROLL = &HFFFFFFFB

Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7

Public Function IsAtBottom(ByVal WB As WebBrowser_V1) As Boolean
    Dim scrlINF As SCROLLINFO
    Dim scrlBarINF As SCROLLBARINFO
    
    With scrlINF
        .cbSize = Len(scrlINF)
        .fMask = SIF_ALL
    End With
    
    GetScrollInfo WB.hwnd, SB_VERT, scrlINF
    
    scrlBarINF.cbSize = Len(scrlBarINF)
    GetScrollBarInfo WB.hwnd, OBJID_VSCROLL, scrlBarINF
    
    IsAtBottom = (scrlINF.nMax - scrlINF.nPos) < (scrlBarINF.rcScrollBar.Bottom - scrlBarINF.rcScrollBar.Top)
End Function

Public Sub MoveToBottom(ByVal WB As WebBrowser_V1)
    SendMessage WB.hwnd, WM_VSCROLL, SB_BOTTOM, 0&
End Sub

