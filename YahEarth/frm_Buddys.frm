VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_Buddys 
   Caption         =   "YahEarth - Buddy List"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Buddys.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView lst_Buddy 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7223
      _Version        =   327682
      Indentation     =   3
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   1
      ImageList       =   "IMG_Icons"
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   120
      Tag             =   "1"
      Top             =   180
      Width           =   645
   End
   Begin ComctlLib.ImageList IMG_Icons 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Buddys.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Buddys.frx":58F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Buddys.frx":5E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Buddys.frx":5F58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image img_ToolBar 
      Height          =   750
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9420
   End
End
Attribute VB_Name = "frm_Buddys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    lst_Buddy.Width = Me.ScaleWidth
    img_ToolBar.Width = Me.Width + 300
    
    lst_Buddy.Height = Me.ScaleHeight - img_ToolBar.Height + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
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
End Sub

Private Sub lst_Buddy_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_Buddy_Click()
    On Error Resume Next
    If lst_Buddy.SelectedItem.Text = "" Then Exit Sub
    If Not Left(lst_Buddy.SelectedItem.Key, 2) = "u_" Then
        If lst_Buddy.SelectedItem.Expanded = True Then
            lst_Buddy.SelectedItem.Expanded = False
        Else
            lst_Buddy.SelectedItem.Expanded = True
        End If
        If lst_Buddy.SelectedItem.Image = 3 Then
            lst_Buddy.SelectedItem.Image = 4
        Else
            lst_Buddy.SelectedItem.Image = 3
        End If
    End If
End Sub

Private Sub lst_Buddy_DblClick()
    On Error Resume Next
    If lst_Buddy.SelectedItem.Text = "" Then Exit Sub
    
    If Mid(lst_Buddy.SelectedItem.Key, 1, 2) = "u_" Then
        FindPm Mid(lst_Buddy.SelectedItem.Key, 3), False, True
    Else
        If lst_Buddy.SelectedItem.Image = 3 Then
            lst_Buddy.SelectedItem.Image = 4
        Else
            lst_Buddy.SelectedItem.Image = 3
        End If
    End If
End Sub

Sub LoadImages()
    img_ToolBar.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\bg.jpg")
    Image1.Picture = LoadPicture(App.Path & "\Resources\Pictures\Buttons\add_mouseOut.gif")
End Sub
