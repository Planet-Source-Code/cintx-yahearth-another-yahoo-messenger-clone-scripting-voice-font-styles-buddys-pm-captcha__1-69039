VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Options"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Apply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Top             =   3960
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.TreeView lst_Options 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7223
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Save 
      Caption         =   "OK"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame frame_Option 
      Caption         =   "Spam Filter"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame frame_Option 
      Caption         =   "Style"
      Height          =   3735
      Index           =   6
      Left            =   2640
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   1100
         Left            =   90
         ScaleHeight     =   1095
         ScaleWidth      =   5505
         TabIndex        =   78
         Top             =   1350
         Width           =   5505
         Begin VB.Frame Frame16 
            Caption         =   "Message Box"
            Height          =   1095
            Left            =   30
            TabIndex        =   79
            Top             =   0
            Width           =   5415
            Begin VB.TextBox txt_MsgBack 
               Height          =   285
               Left            =   1800
               TabIndex        =   85
               Top             =   360
               Width           =   2415
            End
            Begin VB.PictureBox Picture23 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   4320
               ScaleHeight     =   255
               ScaleWidth      =   975
               TabIndex        =   83
               Top             =   360
               Width           =   975
               Begin VB.CommandButton cmd_Pick1 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   0
                  TabIndex        =   84
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.TextBox txt_MsgFont 
               Height          =   285
               Left            =   1800
               TabIndex        =   82
               Top             =   720
               Width           =   2415
            End
            Begin VB.PictureBox Picture9 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   4320
               ScaleHeight     =   255
               ScaleWidth      =   975
               TabIndex        =   80
               Top             =   720
               Width           =   975
               Begin VB.CommandButton cmd_Pick2 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   0
                  TabIndex        =   81
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.Label Label12 
               Caption         =   "Background Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label Label13 
               Caption         =   "Font Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   86
               Top             =   720
               Width           =   1575
            End
         End
      End
      Begin VB.PictureBox Picture10 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   90
         ScaleHeight     =   1185
         ScaleWidth      =   5505
         TabIndex        =   68
         Top             =   2430
         Width           =   5505
         Begin VB.Frame Frame17 
            Caption         =   "User List Box"
            Height          =   1095
            Left            =   30
            TabIndex        =   69
            Top             =   90
            Width           =   5415
            Begin VB.PictureBox Picture22 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   4320
               ScaleHeight     =   255
               ScaleWidth      =   975
               TabIndex        =   74
               Top             =   720
               Width           =   975
               Begin VB.CommandButton cmd_Pick4 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   0
                  TabIndex        =   75
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.TextBox txt_LstFont 
               Height          =   285
               Left            =   1800
               TabIndex        =   73
               Top             =   720
               Width           =   2415
            End
            Begin VB.PictureBox Picture11 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   4320
               ScaleHeight     =   255
               ScaleWidth      =   975
               TabIndex        =   71
               Top             =   360
               Width           =   975
               Begin VB.CommandButton cmd_Pick3 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   0
                  TabIndex        =   72
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.TextBox txt_LstBack 
               Height          =   285
               Left            =   1800
               TabIndex        =   70
               Top             =   360
               Width           =   2415
            End
            Begin VB.Label Label14 
               Caption         =   "Font Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label15 
               Caption         =   "Background Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   360
               Width           =   2655
            End
         End
      End
      Begin VB.Frame Frame15 
         Height          =   975
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   5415
         Begin VB.PictureBox Picture12 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   90
            ScaleHeight     =   375
            ScaleWidth      =   2445
            TabIndex        =   48
            Top             =   180
            Width           =   2445
            Begin VB.CheckBox check_BG 
               Caption         =   "Background Picture"
               Height          =   255
               Left            =   0
               TabIndex        =   49
               Top             =   90
               Width           =   2295
            End
         End
         Begin VB.TextBox txt_Picture 
            Height          =   285
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   4095
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4320
            ScaleHeight     =   255
            ScaleWidth      =   975
            TabIndex        =   44
            Top             =   600
            Width           =   975
            Begin VB.CommandButton cmd_Browse 
               Caption         =   "Browse"
               Height          =   255
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   975
            End
         End
      End
   End
   Begin VB.Frame frame_Option 
      Caption         =   "General"
      Height          =   3735
      Index           =   1
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Frame Frame14 
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   5415
         Begin VB.PictureBox Picture15 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   4065
            TabIndex        =   54
            Top             =   240
            Width           =   4065
            Begin VB.CheckBox check_Reconnect 
               Caption         =   "Automatically reconnect on Disconnects"
               Height          =   255
               Left            =   0
               TabIndex        =   55
               Top             =   0
               Width           =   5175
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5415
         Begin VB.PictureBox Picture13 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   90
            ScaleHeight     =   300
            ScaleWidth      =   2175
            TabIndex        =   50
            Top             =   230
            Width           =   2175
            Begin VB.CheckBox check_Autostart 
               Caption         =   "Start with Windows "
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   5415
         Begin VB.PictureBox Picture14 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   2355
            TabIndex        =   52
            Top             =   240
            Width           =   2355
            Begin VB.CheckBox check_AutoLogin 
               Caption         =   "Auto Login "
               Height          =   255
               Left            =   0
               TabIndex        =   53
               Top             =   0
               Width           =   1455
            End
         End
      End
   End
   Begin VB.Frame frame_Option 
      Caption         =   "Scripting"
      Height          =   3735
      Index           =   4
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5415
         Begin VB.PictureBox Picture16 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   2625
            TabIndex        =   56
            Top             =   230
            Width           =   2625
            Begin VB.CheckBox check_Scripting 
               Caption         =   "Enable Scripting"
               Height          =   255
               Left            =   0
               TabIndex        =   57
               Top             =   0
               Width           =   1575
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5415
         Begin VB.TextBox txt_ScriptTimeout 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Text            =   "10"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Script Timeout (Sec):"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame frame_Option 
      Caption         =   "Chat Room"
      Height          =   3735
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   5415
         Begin VB.PictureBox Picture18 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   2985
            TabIndex        =   60
            Top             =   230
            Width           =   2985
            Begin VB.CheckBox check_BlockDupe 
               Caption         =   "Block duplicate Messages"
               Height          =   255
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5415
         Begin VB.PictureBox Picture17 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   2895
            TabIndex        =   58
            Top             =   230
            Width           =   2895
            Begin VB.CheckBox check_DisableFont 
               Caption         =   "Disable Font Formatting"
               Height          =   255
               Left            =   0
               TabIndex        =   59
               Top             =   0
               Width           =   2175
            End
         End
      End
   End
   Begin VB.Frame frame_Option 
      Caption         =   "Ignore List"
      Height          =   3735
      Index           =   5
      Left            =   2640
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame Frame12 
         Height          =   615
         Left            =   3360
         TabIndex        =   35
         Top             =   240
         Width           =   2175
         Begin VB.PictureBox Picture19 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   1995
            TabIndex        =   62
            Top             =   230
            Width           =   1995
            Begin VB.CheckBox check_IgnoreUser 
               Caption         =   "Enable Ignore"
               Height          =   255
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   1935
            End
         End
      End
      Begin VB.Frame Frame11 
         Height          =   975
         Left            =   3360
         TabIndex        =   31
         Top             =   960
         Width           =   2175
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   33
            Top             =   600
            Width           =   1935
            Begin VB.CommandButton cmd_AddIgnore 
               Caption         =   "Add"
               Height          =   255
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Width           =   1935
            End
         End
         Begin VB.TextBox txt_Add 
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   3135
         Begin VB.ListBox lst_Ignores 
            Height          =   2985
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame13 
         Height          =   975
         Left            =   3360
         TabIndex        =   36
         Top             =   2040
         Width           =   2175
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   39
            Top             =   600
            Width           =   1935
            Begin VB.CommandButton Command1 
               Caption         =   "Clear"
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   0
               Width           =   1935
            End
         End
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   37
            Top             =   240
            Width           =   1935
            Begin VB.CommandButton cmd_RemoveIgnore 
               Caption         =   "Remove"
               Height          =   255
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   1935
            End
         End
      End
   End
   Begin VB.Frame frame_Option 
      Caption         =   "Spam Filter"
      Height          =   3735
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame Frame7 
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   5415
         Begin VB.PictureBox Picture20 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   2895
            TabIndex        =   64
            Top             =   230
            Width           =   2895
            Begin VB.CheckBox check_SpamFilter 
               Caption         =   "Enable Spam Filter"
               Height          =   255
               Left            =   0
               TabIndex        =   65
               Top             =   0
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame8 
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   5415
         Begin VB.PictureBox Picture21 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   5235
            TabIndex        =   66
            Top             =   230
            Width           =   5235
            Begin VB.CheckBox check_SpamUserOnly 
               Caption         =   "Only Check if Message contains the Username"
               Height          =   255
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   5175
            End
         End
      End
      Begin VB.Frame Frame9 
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   5415
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1320
            ScaleHeight     =   255
            ScaleWidth      =   1095
            TabIndex        =   26
            Top             =   1800
            Width           =   1095
            Begin VB.CommandButton cmd_Clear 
               Caption         =   "Clear"
               Height          =   255
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1095
            TabIndex        =   24
            Top             =   1800
            Width           =   1095
            Begin VB.CommandButton cmd_Remove 
               Caption         =   "Remove"
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.ListBox lst_Keywords 
            Height          =   1425
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   5175
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4320
            ScaleHeight     =   255
            ScaleWidth      =   975
            TabIndex        =   22
            Top             =   1800
            Width           =   975
            Begin VB.CommandButton cmd_Add 
               Caption         =   "Add"
               Height          =   255
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.TextBox txt_SpamWord 
            Height          =   285
            Left            =   2520
            TabIndex        =   21
            Top             =   1800
            Width           =   1695
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "YahEarth Version: 1.0"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
End
Attribute VB_Name = "frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub SaveSettings()
    SaveOption "AutoStart", check_Autostart.Value
    SaveOption "AutoLogin", check_AutoLogin.Value
    SaveOption "Scripting", check_Scripting.Value
    SaveOption "ScriptTimeOut", txt_ScriptTimeout
    SaveOption "DisableFonts", check_DisableFont.Value
    SaveOption "BlockDuplicates", check_BlockDupe.Value
    SaveOption "SpamFilter", check_SpamFilter.Value
    SaveOption "SpamUserOnly", check_SpamUserOnly.Value
    SaveFile lst_Keywords, App.Path & "\Resources\Filter\Spam.txt"
    SaveOption "Background", txt_Picture
    SaveOption "MsgBack", txt_MsgBack
    SaveOption "MsgFont", txt_MsgFont
    SaveOption "LstBack", txt_LstBack
    SaveOption "LstFont", txt_LstFont
    SaveOption "Reconnect", check_Reconnect.Value
    SaveOption "UseBG", check_BG.Value
    
    GetSettings 'Reload new Config
End Sub

Sub GetSettings()
    check_Autostart.Value = GetOption("AutoStart", 1)
    check_AutoLogin.Value = GetOption("AutoLogin", 0)
    check_Scripting.Value = GetOption("Scripting", 1)
    check_BlockDupe.Value = GetOption("BlockDuplicates", 1)
    check_DisableFont.Value = GetOption("DisableFonts", 0)
    check_SpamFilter.Value = GetOption("SpamFilter", 1)
    check_SpamUserOnly.Value = GetOption("SpamUserOnly", 1)
    check_Reconnect.Value = GetOption("Reconnect", 1)
    txt_Picture = GetOption("Background", "")
    txt_MsgBack = GetOption("MsgBack", "#FFFFFF")
    txt_MsgFont = GetOption("MsgFont", "#000000")
    txt_LstBack = GetOption("LstBack", "#FFFFFF")
    txt_LstFont = GetOption("LstFont", "#000000")
    check_BG = GetOption("UseBG", 0)
    
    'Enable Functions assign to Options
    Options.blScripting = ValToBool(GetOption("Scripting", 1))
    Options.intScriptTimeOut = GetOption("ScriptTimeOut", 10000)
    Options.blBlockDuplicates = ValToBool(GetOption("BlockDuplicates", 1))
    Options.blDisableFontStyle = ValToBool(GetOption("DisableFonts", 0))
    Options.blSpamFilter = ValToBool(GetOption("SpamFilter", 1))
    Options.blSpamAndUser = ValToBool(GetOption("SpamUserOnly", 1))
    Options.blReconnect = ValToBool(GetOption("Reconnect", 1))
    Options.blUseBG = ValToBool(check_BG.Value)
End Sub

Private Sub cmd_Add_Click()
    lst_Keywords.AddItem txt_SpamWord
End Sub

Private Sub cmd_Apply_Click()
    SaveSettings
    frm_Main.WB.Navigate "about:blank"
    frm_Splash.LoadSettings
End Sub

Private Sub cmd_Browse_Click()
    CMD.Filter = "All Files (*.*) | *.*"
    CMD.ShowOpen
    If Not CMD.filename = "" Then txt_Picture = CMD.filename
    CMD.filename = ""
End Sub

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_Clear_Click()
    lst_Keywords.Clear
End Sub

Private Sub cmd_Pick1_Click()
    frm_Color.LoadColorForm Me, True, txt_MsgBack
    frm_Color.txt_Hex = txt_MsgBack
End Sub

Private Sub cmd_Pick2_Click()
    frm_Color.LoadColorForm Me, True, txt_MsgFont
    frm_Color.txt_Hex = txt_MsgFont
End Sub

Private Sub cmd_Pick3_Click()
    frm_Color.LoadColorForm Me, True, txt_LstBack
    frm_Color.txt_Hex = txt_LstBack
End Sub

Private Sub cmd_Pick4_Click()
    frm_Color.LoadColorForm Me, True, txt_LstFont
    frm_Color.txt_Hex = txt_LstFont
End Sub

Private Sub cmd_Remove_Click()
    lst_Keywords.RemoveItem lst_Keywords.ListIndex
    cmd_Remove.Enabled = False
End Sub

Private Sub cmd_Save_Click()
    SaveSettings
    frm_Main.WB.Navigate "about:blank"
    frm_Splash.LoadSettings
    Unload Me
End Sub


Private Sub Form_Load()
    lst_Options.Nodes.Add , , "General", "General"
    lst_Options.Nodes.Add , , "Spam Filter", "Spam Filter"
    lst_Options.Nodes.Add , , "Chat Room", "Chat Room"
    lst_Options.Nodes.Add , , "Scripting", "Scripting"
    lst_Options.Nodes.Add , , "Ignore List", "Ignore List"
    lst_Options.Nodes.Add , , "Style", "Style"
    
    OpenFile lst_Keywords, App.Path & "\Resources\Filter\Spam.txt"
    GetSettings
End Sub

Private Sub lst_Keywords_Click()
    On Error Resume Next
    If lst_Keywords = "" Then
        cmd_Remove.Enabled = False
    Else
        cmd_Remove.Enabled = True
    End If
End Sub

Private Sub lst_Options_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_Options_Click()
    Dim X As Integer
    On Error Resume Next
    For X = 1 To lst_Options.Nodes.Count
        frame_Option(X).Visible = False
    Next X
    frame_Option(lst_Options.SelectedItem.Index).Visible = True
End Sub
