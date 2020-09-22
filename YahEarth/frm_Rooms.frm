VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_Rooms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo! Chat Rooms"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Rooms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_Refresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   4800
      Width           =   855
   End
   Begin ComctlLib.TreeView lst_Cats 
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8070
      _Version        =   327682
      Indentation     =   531
      Style           =   1
      ImageList       =   "img_List"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.TreeView lst_Rooms 
      Height          =   4335
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7646
      _Version        =   327682
      Indentation     =   3
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txt_Room 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   6615
   End
   Begin ComctlLib.TreeView lst_URooms 
      Height          =   4335
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7646
      _Version        =   327682
      Indentation     =   3
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd_Join 
      Caption         =   "Join"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   4800
      Width           =   735
   End
   Begin ComctlLib.ImageList img_List 
      Left            =   840
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rooms.frx":57E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label bt_URooms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " User Room's"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label bt_Rooms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Room's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_Rooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data_Cats As String
Dim Data_Rooms As String
Dim strCat As String

Private Sub bt_Rooms_Click()
    bt_Rooms.FontBold = True
    lst_Rooms.Visible = True
    lst_URooms.Visible = False
    bt_URooms.FontBold = False
End Sub

Private Sub bt_URooms_Click()
    bt_Rooms.FontBold = False
    lst_Rooms.Visible = False
    lst_URooms.Visible = True
    bt_URooms.FontBold = True
End Sub

Sub StartBrowse(FromLogin As Boolean)
    Dim strData As String
    
    If FromLogin = True Then
        cmd_OK.Visible = True
        cmd_Join.Visible = False
    Else
        cmd_Join.Visible = True
        cmd_OK.Visible = False
    End If
        
    Me.Show
    BrowseCats
End Sub

Sub BrowseCats()
    Dim strData As String
    lst_Cats.Nodes.Clear
    lst_Cats.Nodes.Add , , "loadingcaption", "Loading..."
    If Inet1.StillExecuting = True Then
        Inet1.Cancel
        If Inet1.StillExecuting = True Then
            Inet1.Cancel
            If Inet1.StillExecuting = True Then
                Exit Sub
            End If
        End If
    End If
    strData = Inet1.OpenUrl("http://insider.msg.yahoo.com/ycontent/?chatcat")
    lst_Cats.Nodes.Clear
    GetCats strData
End Sub

Sub GetCats(strData As String)
    Dim X As Integer
    Dim strCats() As String
    Dim strCat As String
    Dim strSub() As String
    Dim intCount As Integer
    Dim i As Integer
        
    strCats = Split(strData, "<category", , vbTextCompare)
    For X = 1 To UBound(strCats)
        strCat = Parse("name=" & Chr(34), Chr(34), strCats(X))
        strCat = Replace(strCat, "&amp;", "&")
        strID = Parse("id=" & Chr(34), Chr(34), strCats(X))
        lst_Cats.Nodes.Add , , "chatroom_" & strID, strCat, 1, 1
    Next X
End Sub

Sub GetRooms(strData As String)
    Dim strType As String
    Dim strRoom As String
    Dim strID As String
    Dim strLobby() As String
    Dim strRooms() As String
    Dim X As Integer
    Dim i As Integer
        
    strRooms = Split(strData, "<room", , vbTextCompare)
    For X = 1 To UBound(strRooms)
        strType = Parse("type=" & Chr(34), Chr(34), strRooms(X))
        If LCase(strType) = "yahoo" Then
            strID = Parse("id=" & Chr(34), Chr(34), strRooms(X))
            strRoom = Parse("name=" & Chr(34), Chr(34), strRooms(X))
            strLobby = Split(strRooms(X), "<lobby")
            If UBound(strLobby) > 2 Then
                lst_Rooms.Nodes.Add , , "cat_" & strID, strRoom
                For i = 1 To UBound(strLobby)
                    lst_Rooms.Nodes.Add "cat_" & strID, tvwChild, "room_" & strID & "_" & i, strRoom & ":" & i
                Next i
            Else
                lst_Rooms.Nodes.Add , , "room_" & strID, strRoom & ":1"
            End If
        ElseIf LCase(strType) = "user" Then
            strID = Parse("id=" & Chr(34), Chr(34), strRooms(X))
            strRoom = Parse("name=" & Chr(34), Chr(34), strRooms(X))
            lst_URooms.Nodes.Add , , "room_" & strID, strRoom
        End If
    Next X
End Sub

Sub LoadRoom(strCat As String)
    Dim strData As String
    
    On Error Resume Next
    
    lst_Rooms.Nodes.Clear
    lst_URooms.Nodes.Clear
    lst_Rooms.Nodes.Add , , "loadingcaption", "Loading..."
    lst_URooms.Nodes.Add , , "loadingcaption", "Loading..."
    If Inet1.StillExecuting = True Then
        Inet1.Cancel
        If Inet1.StillExecuting = True Then
            Inet1.Cancel
            If Inet1.StillExecuting = True Then
                Exit Sub
            End If
        End If
    End If
    strData = Inet1.OpenUrl("http://insider.msg.yahoo.com/ycontent/?" & strCat)
    lst_Rooms.Nodes.Clear
    lst_URooms.Nodes.Clear
    GetRooms strData
End Sub

Private Sub cmd_Join_Click()
    If Not txt_Room = "" Then
        DoCommand "/join " & txt_Room
        Unload Me
    End If
End Sub

Private Sub cmd_OK_Click()
    If txt_Room = "" Then Exit Sub
    frm_Login.txt_Room = txt_Room
    frm_Login.SetFocus
    Unload Me
End Sub

Private Sub cmd_Refresh_Click()
    BrowseCats
End Sub

Private Sub Form_Load()
    InitCommonControls
    ExecuteScript 7, , , , Me.Name
End Sub

Private Sub lst_Cats_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_Cats_Click()
    On Error Resume Next
    LoadRoom lst_Cats.SelectedItem.Key
End Sub

Private Sub lst_Rooms_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_Rooms_Click()
    If Left(lst_Rooms.SelectedItem.Key, 4) = "room" Then
        txt_Room = lst_Rooms.SelectedItem.Text
    End If
End Sub

Private Sub lst_Rooms_DblClick()
    If cmd_Join.Visible = True Then
        If txt_Room = "" Then Exit Sub
        DoCommand "/join " & txt_Room
        Unload Me
    Else
        If txt_Room = "" Then Exit Sub
        frm_Login.txt_Room = txt_Room
        frm_Login.SetFocus
        Unload Me
    End If
End Sub
