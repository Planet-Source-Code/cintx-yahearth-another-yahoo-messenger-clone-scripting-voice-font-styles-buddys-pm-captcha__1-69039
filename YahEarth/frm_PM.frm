VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_PM 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   8145
   ClientTop       =   6150
   ClientWidth     =   6645
   Icon            =   "frm_PM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tlb_Buttons 
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   741
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      ImageList       =   "img_Icons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      Begin VB.ComboBox cmb_Font 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   30
         Width           =   2535
      End
      Begin VB.ComboBox cmb_Size 
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   30
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3345
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd_Send 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      ExtentX         =   11245
      ExtentY         =   3836
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox txt_Message 
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   979
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frm_PM.frx":57E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Buffer 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frm_PM.frx":5859
   End
   Begin ComctlLib.ImageList img_Icons 
      Left            =   6600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":58DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":5E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":63E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":64F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":660B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":671D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":6C6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_PM.frx":7229
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
   End
End
Attribute VB_Name = "frm_PM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLastSend As Double
Sub InitWindow()
    InitCommonControls
    WB.Navigate "about:blank"
End Sub

Sub LoadWindow()
    ToolFlat tlb_Buttons, True
    ToolFlat cmd_Send, True
    ColorMenu Me, &HEDEFEF
End Sub

Sub LoadFonts()
    Dim X As Integer
    For X = 0 To Screen.FontCount - 1
        cmb_Font.AddItem Screen.Fonts(X)
    Next X
    For X = 2 To 32 Step 2
        cmb_Size.AddItem X
    Next X
End Sub

Sub LoadImages()
    'img_ToolBar.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\bg.jpg")
    'Image1.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\add_mouseOut.gif")
End Sub

Private Sub cmd_Send_Click()
    Dim strMsg As String
    If Timer - strLastSend < 0.5 Then
        MsgBox "You are sending too fast Messages"
        Exit Sub
    Else
        strLastSend = Timer
    End If
    Buffer = txt_Message
    txt_Message.Text = ""
    strMsg = GenerateHTML(Buffer)
    ProcessHTML YMSG.strUser, strMsg, WB, True
    If Not InStr(PMi(Me.Tag).strTo, "@hotmail") = 0 Then
        SendData SendPM(YMSG.strUser, PMi(Me.Tag).strTo, "<font INF ID:YEH Proto:YMSG>" & strMsg, True)
    Else
        SendData SendPM(YMSG.strUser, PMi(Me.Tag).strTo, "<font INF ID:YEH Proto:YMSG>" & strMsg, False)
    End If
End Sub

Private Sub Form_Load()
    WB.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 6075 Then
        Me.Width = 6075
    End If
    
    If Me.Height < 4215 Then
        Me.Height = 4215
    End If
    
    WB.Width = Me.Width - 150
    tlb_Buttons.Width = Me.Width
    txt_Message.Width = Me.Width - cmd_Send.Width - 150
    cmd_Send.Left = txt_Message.Width + 10
    StatusBar1.Panels(1).Width = Me.Width
    
    WB.Height = Me.Height - 1910
    lst_User.Height = Me.Height - 2710
    tlb_Buttons.Top = WB.Height
    cmd_Send.Top = tlb_Buttons.Top + 430
    txt_Message.Top = cmd_Send.Top - 40
    lbl_Voice.Width = Me.Width - 7400
    bg_Voice.Width = Me.Width
    StatusBar1.Panels(1).Width = Me.ScaleWidth
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMi(Me.Tag).blUsed = False
    PMi(Me.Tag).strTo = ""
    intPmCount = intPmCount - 1
    SaveSetting "YahEarth", "PM" & Me.Tag, "Top", Me.Top
    SaveSetting "YahEarth", "PM" & Me.Tag, "Left", Me.Left
    Unload Me
End Sub

Private Sub Image1_Click()
    frm_Add.Show
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image1.Tag = 1 Then
        Image1.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\add_mouseOver.gif")
        Image1.Tag = 2
    End If
End Sub

Private Sub img_ToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image1.Tag = 2 Then
        Image1.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\add_mouseOut.gif")
        Image1.Tag = 1
    End If
End Sub

Private Sub tlb_Buttons_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1
            frm_Smileys.OpenSmileys Me
            If (Screen.Height) - ((Me.Top + Me.ScaleHeight) + 500) < frm_Smileys.ScaleHeight Then
                frm_Smileys.Top = Me.Top + tlb_Buttons.Top - 1600
            Else
                frm_Smileys.Top = Me.Top + tlb_Buttons.Top + 1100
            End If
            If (Screen.Width) - ((Me.Left) + 500) < frm_Smileys.ScaleWidth Then
                frm_Smileys.Left = Me.Left - frm_Smileys.ScaleWidth
            Else
                frm_Smileys.Left = Me.Left + 400
            End If
        Case 2
            DoPMVoice Me.Tag
        Case 4
            If Button.Value = tbrPressed Then
                txt_Message.SelBold = True
            Else
                txt_Message.SelBold = False
            End If
        Case 5
            If Button.Value = tbrPressed Then
                txt_Message.SelItalic = True
            Else
                txt_Message.SelItalic = False
            End If
        Case 6
            If Button.Value = tbrPressed Then
                txt_Message.SelUnderline = True
            Else
                txt_Message.SelUnderline = False
            End If
        Case 7
            frm_Color.LoadColorForm Me
            frm_Color.txt_Hex = RGB2Hex(txt_Message.SelColor)
            frm_Color.Hex2RGB_2 frm_Color.txt_Hex
    End Select
End Sub

Private Sub txt_Message_Change()
    If Not Trim(txt_Message.Text) = "" Then
        cmd_Send.Enabled = True
        If Len(Trim(txt_Message.Text)) = 1 Then
            SendData Typing(YMSG.strUser, PMi(Me.Tag).strTo)
        End If
    Else
        cmd_Send.Enabled = False
    End If
End Sub

Private Sub txt_Message_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmd_Send.Enabled = True Then cmd_Send_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txt_Message_SelChange()
    On Error Resume Next
    
    cmb_Font.Text = txt_Message.SelFontName
    cmb_Size.Text = Val(txt_Message.SelFontSize)
    If txt_Message.SelBold = True Then
        tlb_Buttons.Buttons(4).Value = tbrPressed
    Else
        tlb_Buttons.Buttons(4).Value = tbrUnpressed
    End If
    If txt_Message.SelItalic = True Then
        tlb_Buttons.Buttons(5).Value = tbrPressed
    Else
        tlb_Buttons.Buttons(5).Value = tbrUnpressed
    End If
    If txt_Message.SelUnderline = True Then
        tlb_Buttons.Buttons(6).Value = tbrPressed
    Else
        tlb_Buttons.Buttons(6).Value = tbrUnpressed
    End If
End Sub

Private Sub WB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If Not URL = "about:blank" Then
        If Left(LCase(URL), "7") = "http://" Then
            OpenUrl URL
        End If
        Cancel = True
    End If
End Sub

Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    WB.Document.write "<html>" & vbCrLf & frm_Main.txt_NoClick
    WB.Visible = True
End Sub
