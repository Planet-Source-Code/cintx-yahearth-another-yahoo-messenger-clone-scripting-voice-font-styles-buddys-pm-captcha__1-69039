VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_Reconnect 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   360
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17637
            MinWidth        =   17637
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   90
         ScaleHeight     =   285
         ScaleWidth      =   3075
         TabIndex        =   20
         Top             =   2970
         Width           =   3075
         Begin VB.CheckBox check_Join 
            Caption         =   "Join Chat after Login"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   90
         ScaleHeight     =   285
         ScaleWidth      =   3075
         TabIndex        =   18
         Top             =   2610
         Width           =   3075
         Begin VB.CheckBox check_Voice 
            Caption         =   "Turn Voice On"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   2895
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   90
         ScaleHeight     =   285
         ScaleWidth      =   3075
         TabIndex        =   16
         Top             =   2250
         Width           =   3075
         Begin VB.CheckBox check_Save 
            Caption         =   "Save Username && Password"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.ComboBox Protocol 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1800
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   1090
         Width           =   255
         Begin VB.CommandButton cmd_Browse 
            Caption         =   ".."
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txt_Room 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   10
         Top             =   3360
         Width           =   1935
         Begin VB.CommandButton cmd_Login 
            Caption         =   "Login"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.ComboBox Server 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "scsc.msg.yahoo.com:5050"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txt_Pass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt_User 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Protocol:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Room:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Browse_Click()
    frm_Rooms.StartBrowse True
End Sub

Private Sub cmd_Login_Click()
    Login
End Sub

Sub Login()
    Dim Host(1) As String
    
    SaveSettings
    
    Select Case Mid(Protocol, 1, 6)
        Case "YMSG12"
            YMSG_VER = 12
        Case "YMSG15"
            YMSG_VER = 15
        Case Else
            YMSG_VER = 12
    End Select
    
    If (Server = "") Then
        Host(0) = "scs.msg.yahoo.com"
        Host(1) = "5050"
    Else
        Host(0) = Split(Server, ":")(0)
        Host(1) = Split(Server, ":")(1)
        If (Host(1) = "") Then
            Host(1) = "5050"
        End If
    End If
    
    If Not Host(0) = "" And Not Host(1) = "" Then
        YMSG.strUser = txt_User
        YMSG.strPass = txt_Pass
        YMSG.strRoom = txt_Room
        StatusBar1.Panels(1).Text = "Status: Connecting"
        Socket.Close
        Socket.Connect Host(0), Host(1)
    End If
End Sub

Private Sub Form_Activate()
    InitCommonControls
    StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_Load()
    InitCommonControls
    frm_Main.Show
    frm_Splash.Show
    Me.Visible = False
    LoadSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
    frm_Main.Enabled = True
    frm_Main.SetFocus
End Sub

Sub LoadSettings()
    On Error Resume Next
    Protocol.AddItem "YMSG12"
    Protocol.AddItem "YMSG13"
    Protocol.AddItem "YMSG14"
    Protocol.AddItem "YMSG15 (Beta)"
    txt_User = GetSetting("YahEarth", "Login", "User", "")
    txt_Pass = GetSetting("YahEarth", "Login", "Pass", "")
    check_Save.Value = GetSetting("YahEarth", "Login", "Save", 0)
    txt_Room = GetSetting("YahEarth", "Login", "Room", "Yahoo! Chat Help:1")
    check_Voice.Value = GetSetting("YahEarth", "Login", "Voice", 1)
    check_Join.Value = GetSetting("YahEarth", "Login", "Join", 1)
    Protocol = GetSetting("YahEarth", "Login", "Protocol", "YMSG12")
End Sub

Sub SaveSettings()
    If check_Save.Value = 1 Then
        SaveSetting "YahEarth", "Login", "User", txt_User
        SaveSetting "YahEarth", "Login", "Pass", txt_Pass
    End If
    SaveSetting "YahEarth", "Login", "Save", check_Save.Value
    SaveSetting "YahEarth", "Login", "Room", txt_Room
    SaveSetting "YahEarth", "Login", "Voice", check_Voice.Value
    SaveSetting "YahEarth", "Login", "Join", check_Join.Value
    SaveSetting "YahEarth", "Login", "Protocol", Protocol
End Sub

Private Sub Socket_Close()
    YMSG.blJoined = False
    YMSG.strKey = ""
    YMSG.strRoom = ""
    YMSG.strRoomSpace = ""
    YMSG.strVoiceKey = ""
    blVoice = False
    Status 2, "Online Status: Offline"
    
    If Options.blReconnect = True Then
        ProcessError frm_Main.WB, "Connection Lost, reconnecting in 3 Seconds"
        Timer_Reconnect = True
    Else
        ProcessError frm_Main.WB, "Connection Lost"
    End If
End Sub

Private Sub Socket_Connect()
    StatusBar1.Panels(1).Text = "Status: Connected"
    SendData GetHash(YMSG.strUser)
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
            
    Socket.GetData strData, vbString, bytesTotal
    
    IncommingData strData
End Sub

Private Sub Socket_SendComplete()
    blSend = False
End Sub

Private Sub Timer_Reconnect_Timer()
    Login
    Timer_Reconnect = False
End Sub
