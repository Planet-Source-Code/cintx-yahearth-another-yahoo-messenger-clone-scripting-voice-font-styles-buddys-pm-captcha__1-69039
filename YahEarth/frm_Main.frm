VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{2B323CCC-50E3-11D3-9466-00A0C9700498}#1.0#0"; "yacscom.dll"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_Main 
   Caption         =   "YahEarth"
   ClientHeight    =   5610
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9585
   DrawMode        =   11  'Not Xor Pen
   DrawStyle       =   1  'Dash
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tlb_Buttons 
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   741
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      ImageList       =   "img_Icons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   3460
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      Begin VB.PictureBox bg_Voice 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6100
         ScaleHeight     =   300
         ScaleWidth      =   3255
         TabIndex        =   9
         Top             =   40
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CheckBox check_Talk 
            Caption         =   "Talk"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   30
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmd_Talk 
            Caption         =   "Talk"
            CausesValidation=   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   30
            Width           =   735
         End
         Begin VB.CommandButton cmd_Auto 
            Caption         =   "A"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   30
            Width           =   255
         End
         Begin VB.Image pic_VoiceIn 
            Height          =   30
            Left            =   1080
            Picture         =   "frm_Main.frx":57E2
            Top             =   255
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Image pic_Voice 
            Height          =   30
            Left            =   1080
            Picture         =   "frm_Main.frx":AD0F
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lbl_Voice 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   15
            Width           =   2055
         End
      End
      Begin VB.ComboBox cmb_Size 
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Top             =   30
         Width           =   735
      End
      Begin VB.ComboBox cmb_Font 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   30
         Width           =   2535
      End
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   3615
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   6376
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
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmd_Send 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   4680
      Width           =   855
   End
   Begin RichTextLib.RichTextBox txt_Message 
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   979
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frm_Main.frx":1023C
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
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   7560
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin ComctlLib.ListView lst_User 
      Height          =   3615
      Left            =   7200
      TabIndex        =   3
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   6376
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      Icons           =   "img_List"
      SmallIcons      =   "img_List"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5355
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   351
            MinWidth        =   351
            Text            =   "Status:"
            TextSave        =   "Status:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Online Status: Offline"
            TextSave        =   "Online Status: Offline"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Blocked Spam: 0"
            TextSave        =   "Blocked Spam: 0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Messages: 0"
            TextSave        =   "Messages: 0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Buffer 
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frm_Main.frx":102B3
   End
   Begin VB.TextBox txt_NoClick 
      Height          =   3645
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm_Main.frx":1032E
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
   End
   Begin ComctlLib.ImageList img_List 
      Left            =   8160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":10758
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":1085A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":1095C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin YACSCOMLibCtl.YAcs Voice 
      Left            =   2400
      OleObjectBlob   =   "frm_Main.frx":10A5E
      Top             =   120
   End
   Begin ComctlLib.ImageList img_Icons 
      Left            =   8760
      Top             =   120
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
            Picture         =   "frm_Main.frx":10A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":1103C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":1158E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":116A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":117B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":118C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":11E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":123D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   960
      Tag             =   "1"
      Top             =   180
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   120
      Tag             =   "1"
      Top             =   180
      Width           =   645
   End
   Begin VB.Image img_ToolBar 
      Height          =   750
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9420
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_Login 
         Caption         =   "Login"
      End
      Begin VB.Menu mnu_Logout 
         Caption         =   "Logout"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu1_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "Tools"
      Begin VB.Menu mnu_Scripts 
         Caption         =   "Scripting"
      End
      Begin VB.Menu mnu2_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnu_Buddys 
      Caption         =   "Buddys"
      Begin VB.Menu mnu_Buddylist 
         Caption         =   "Buddy List"
      End
      Begin VB.Menu mnu_3_line_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_newPM 
         Caption         =   "New PM"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Help"
      Begin VB.Menu mnu_About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLastSend As Double

Private Sub check_Talk_Click()
    If check_Talk.Value = 1 Then
        Voice.startTransmit
    Else
        Voice.stopTransmit
    End If
End Sub

Private Sub cmb_Font_Click()
    txt_Message.SelFontName = cmb_Font
    SaveOption "FontFace", cmb_Font
End Sub

Private Sub cmb_Size_Click()
    txt_Message.SelFontSize = cmb_Size
    SaveOption "FontSize", cmb_Size
End Sub

Private Sub cmd_Auto_Click()
    If cmd_Talk.Visible = True Then
        cmd_Talk.Visible = False
        check_Talk.Visible = True
    Else
        cmd_Talk.Visible = True
        Voice.stopTransmit
        check_Talk.Value = 0
        check_Talk.Visible = False
    End If
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
    If Left(Buffer.Text, 1) = "/" Then
        DoCommand (Buffer.Text)
        Exit Sub
    End If
    strMsg = GenerateHTML(Buffer)
    SendData SendChat(YMSG.strUser, YMSG.strRoom, "<font INF ID:YEH Proto:YMSG>" & strMsg)
    ProcessHTML YMSG.strUser, strMsg, WB
End Sub

Private Sub cmd_Talk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Voice.startTransmit
End Sub

Private Sub cmd_Talk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Voice.stopTransmit
End Sub

Private Sub Form_Load()
    WB.Navigate "about:blank"
    ExecuteScript 7, , , , Me.Name
    Me.Show
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 8325 Then
        Me.Width = 8325
    End If
    
    If Me.Height < 5700 Then
        Me.Height = 5700
    End If
    
    img_ToolBar.Width = Me.Width + 300
    WB.Width = Me.Width - lst_User.Width - 150
    lst_User.Left = WB.Width + 10
    tlb_Buttons.Width = Me.Width
    txt_Message.Width = Me.Width - cmd_Send.Width - 150
    cmd_Send.Left = txt_Message.Width + 10
    StatusBar1.Panels(1).Width = Me.Width
    
    WB.Height = Me.ScaleHeight - 1900
    lst_User.Height = Me.ScaleHeight - 1900
    tlb_Buttons.Top = lst_User.Height + 710
    cmd_Send.Top = tlb_Buttons.Top + 430
    txt_Message.Top = cmd_Send.Top - 40
    lbl_Voice.Width = Me.Width - 7400
    bg_Voice.Width = Me.Width
    StatusBar1.Panels(1).Width = Me.ScaleWidth - 7000
    StatusBar1.Panels(2).Width = 3000
    StatusBar1.Panels(3).Width = 2000
    StatusBar1.Panels(4).Width = 2000
End Sub

Sub InitWindow()
    InitCommonControls
    txt_Message.SelFontName = GetOption("FontFace", "Arial")
    txt_Message.SelFontSize = GetOption("FontSize", "10")
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
    img_ToolBar.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\bg.jpg")
    Image1.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\add_mouseOut.gif")
    Image2.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\buddy_mouseOut.gif")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Box As String
    Box = MsgBox("Quit YahEarth?", vbYesNo + vbQuestion, "Quit")
    Select Case Box
        Case vbYes
            ExecuteScript 6
            End
        Case vbNo
            Cancel = 1
    End Select
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

Private Sub Image2_Click()
    frm_Rooms.StartBrowse False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image2.Tag = 1 Then
        Image2.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\buddy_mouseOver.gif")
        Image2.Tag = 2
    End If
End Sub

Private Sub img_ToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image1.Tag = 2 Then
        Image1.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\add_mouseOut.gif")
        Image1.Tag = 1
    End If
    If Image2.Tag = 2 Then
        Image2.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\buddy_mouseOut.gif")
        Image2.Tag = 1
    End If
End Sub

Private Sub timer_Send_Timer()
    timer_Send = False
End Sub

Private Sub lst_User_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_User_DblClick()
    On Error Resume Next
    If Not lst_User.SelectedItem.Text = "" Then FindPm lst_User.SelectedItem.Text, , True
End Sub

Private Sub lst_User_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lst_User.SelectedItem.Text = "" Then Exit Sub
    If Button = 2 Then
        If IsIgnored(lst_User.SelectedItem.Text) = True Then
            frm_Menu.mnu_Ignore.Caption = "UnIgnore User"
        Else
            frm_Menu.mnu_Ignore.Caption = "Ignore User"
        End If
        If lst_User.SelectedItem.SmallIcon = 3 Then
            frm_Menu.mnu_Mute.Caption = "UnMute User"
            frm_Menu.mnu_Mute.Enabled = True
        ElseIf lst_User.SelectedItem.SmallIcon = 2 Then
            frm_Menu.mnu_Mute.Caption = "Mute User"
            frm_Menu.mnu_Mute.Enabled = True
        ElseIf lst_User.SelectedItem.SmallIcon = 1 Then
            frm_Menu.mnu_Mute.Caption = "Mute User"
            frm_Menu.mnu_Mute.Enabled = False
        End If
        PopupMenu frm_Menu.menu_List
    End If
End Sub

Private Sub mnu_About_Click()
    frm_About.Show
End Sub

Private Sub mnu_Buddylist_Click()
    frm_Buddys.Show
End Sub

Private Sub mnu_Login_Click()
    frm_Login.Show
End Sub

Private Sub mnu_Logout_Click()
    frm_Login.Socket.Close
    ClearFields
    frm_Login.Show
    Me.mnu_Logout.Enabled = False
    Me.mnu_Login.Enabled = True
End Sub

Private Sub mnu_newPM_Click()
    frm_NewPM.Show
End Sub

Private Sub mnu_Options_Click()
    frm_Options.Show
End Sub

Private Sub mnu_Scripts_Click()
    frm_Scripting.Show
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
    If Panel.Index = 2 Then
        PopupMenu frm_Menu.mnu_Status, , StatusBar1.Panels(1).Width, Me.ScaleHeight
    End If
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As ComctlLib.Panel)
    If Not Panel.Index = 1 And Not Panel.Index = 2 Then
        SendChat YMSG.strUser, YMSG.strRoom, Panel.Text
        ProcessHTML YMSG.strUser, Panel.Text, WB
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
            If Not YMSG.strVoiceKey = "" Then
                DoVoice
            Else
                Button.Value = tbrUnpressed
            End If
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
            'frm_Color.txt_Hex = RGB2Hex(txt_Message.SelColor)
            frm_Color.Hex2RGB_2 frm_Color.txt_Hex
    End Select
End Sub

Private Sub txt_Message_Change()
    If Not Trim(txt_Message.Text) = "" Then
        cmd_Send.Enabled = True
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

Private Sub Voice_onConferenceNotReady()
    If blVoice = True Then
        Status 1, "Status: Voice Disabled"
        blVoice = False
    Else
        Status 1, "Status: Voice Error"
    End If
    tlb_Buttons.Buttons(2).Value = tbrUnpressed
    bg_Voice.Visible = False
End Sub

Private Sub Voice_onConferenceReady()
    Status 1, "Status: Voice Enabled"
    tlb_Buttons.Buttons(2).Value = tbrPressed
    blVoice = True
    bg_Voice.Visible = True
    pic_Voice.Visible = False
    pic_VoiceIn.Visible = False
End Sub

Private Sub Voice_onInputLevelChange(ByVal level As Integer)
    Dim strVal As Integer
    If level < 0 Then
        strVal = Mid(level, 2)
    Else
        strVal = level
    End If
    pic_VoiceIn.Visible = True
    pic_VoiceIn.Width = (strVal * 100) + 155
    DoEvents
End Sub

Private Sub Voice_onLocalOffAir()
    pic_VoiceIn.Visible = False
    lbl_Voice.Caption = ""
End Sub

Private Sub Voice_onLocalOnAir()
    lbl_Voice.Caption = YMSG.strUser
End Sub

Private Sub Voice_onOutputLevelChange(ByVal level As Integer)
    Dim strVal As Integer
    If level < 0 Then
        strVal = Mid(level, 2)
    Else
        strVal = level
    End If
    pic_Voice.Visible = True
    pic_Voice.Width = (strVal * 100) + 155
    DoEvents
End Sub

Private Sub Voice_onRemoteSourceOffAir(ByVal sourceId As Long, ByVal sourceName As String)
    pic_Voice.Visible = False
    lbl_Voice.Caption = ""
End Sub

Private Sub Voice_onRemoteSourceOnAir(ByVal sourceId As Long, ByVal sourceName As String)
    lbl_Voice.Caption = sourceName
End Sub

Private Sub Voice_onSourceEntry(ByVal sourceId As Long, ByVal sourceName As String)
    Dim X As Integer
    For X = 1 To lst_User.ListItems.Count
        If lst_User.ListItems(X).Text = sourceName Then
            lst_User.ListItems(X).SmallIcon = 2
            lst_User.ListItems(X).Tag = sourceId
            Exit For
        End If
    Next X
End Sub

Private Sub Voice_onSourceExit(ByVal sourceId As Long, ByVal sourceName As String)
    Dim X As Integer
    For X = 1 To lst_User.ListItems.Count
        If lst_User.ListItems(X).Text = sourceName Then
            lst_User.ListItems(X).SmallIcon = 1
            lst_User.ListItems(X).Tag = ""
            Exit For
        End If
    Next X
End Sub

Private Sub Voice_onSystemConnectFailure(ByVal code As Long, ByVal message As String)
    Status 1, "Status: Voice Connection Error"
    tlb_Buttons.Buttons(2).Value = tbrUnpressed
End Sub

Private Sub WB_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim strUser As String
    If Not URL = "about:blank" Then
        If Left(LCase(URL), "7") = "http://" Then
            OpenUrl URL
        End If
        If Left(LCase(URL), 9) = "yahearth:" Then
            strUser = Mid(URL, 10)
            FindPm strUser
        End If
        Cancel = True
    End If
End Sub

Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Dim strBG As String
    
    If ValToBool(GetOption("UseBG", 0)) = True Then
        strBG = "<body background=" & Chr(34) & "file://" & GetOption("Background", "") & Chr(34) & " bgproperties=" & Chr(34) & "fixed" & Chr(34) & ">"
    Else
        strBG = ""
    End If
    Me.Visible = True
    WB.Document.write "<html>" & vbCrLf & txt_NoClick & vbCrLf & strBG & strBuffer
    
    WB.Visible = True
End Sub

Private Sub WB_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Cancel = True
End Sub
