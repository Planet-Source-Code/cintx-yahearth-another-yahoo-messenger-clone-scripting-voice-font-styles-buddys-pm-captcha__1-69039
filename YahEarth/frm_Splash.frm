VERSION 5.00
Begin VB.Form frm_Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "frm_Splash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frm_Splash.frx":57E2
   ScaleHeight     =   2160
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer_Start 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lbl_Load 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1890
      Width           =   3705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By cIntX / CiPH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1035
   End
End
Attribute VB_Name = "frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Resources\Pictures\Background\yahearth_logo.jpg")
End Sub

Sub LoadSettings()
    With frm_Main
        .txt_Message.BackColor = Hex2RGB(GetOption("MsgBack", "#FFFFFF"))
        .txt_Message.SelColor = Hex2RGB(GetOption("MsgFont", "#000000"))
        .lst_User.BackColor = Hex2RGB(GetOption("LstBack", "#FFFFFF"))
        .lst_User.ForeColor = Hex2RGB(GetOption("LstFont", "#000000"))
    End With
End Sub

Private Sub timer_Start_Timer()
    Me.Show
    lbl_Load.Caption = "Scripts"
    LoadScripts
    lbl_Load.Caption = "Settings"
    LoadSettings
    frm_Options.GetSettings
    lbl_Load.Caption = "Styles"
    ExecuteScript 5
    frm_Main.InitWindow
    frm_PM.InitWindow
    frm_Buddys.LoadImages
    frm_Main.LoadFonts
    frm_PM.LoadFonts
    frm_Main.LoadImages
    frm_PM.LoadImages
    frm_Main.LoadWindow
    frm_PM.LoadWindow
    Pause 0.1
    Unload Me
    frm_Login.Show
End Sub

Sub LoadScripts()
    blScripting = GetSetting("YahEarth", "Settings", "Scripting", True)
    Src(1).strFile = "IncommingData.vbs"
    Src(2).strFile = "OutgoingData.vbs"
    Src(3).strFile = "IncommingChatText.vbs"
    Src(4).strFile = "IncommingPM.vbs"
    Src(5).strFile = "AppStart.vbs"
    Src(6).strFile = "AppEnd.vbs"
    Src(7).strFile = "NewWindow.vbs"
    Src(8).strFile = "Custom.vbs"
    For X = 1 To UBound(Src)
        Src(X).strScript = LoadTextFile(App.Path & "\Resources\Scripts\" & Src(X).strFile)
    Next X
End Sub
