VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_NewPM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " New PM"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_NewPM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_New 
      Caption         =   "New"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin ComctlLib.ListView lst_Buddy 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      Icons           =   "img_Icons"
      SmallIcons      =   "img_Icons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   5293
      EndProperty
   End
   Begin VB.TextBox txt_To 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin ComctlLib.ImageList img_Icons 
      Left            =   120
      Top             =   4680
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
            Picture         =   "frm_NewPM.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_NewPM.frx":58F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_NewPM.frx":5E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_NewPM.frx":5F58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frm_NewPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_New_Click()
    If Not Trim(txt_To) = "" Then
        FindPm txt_To, False, True
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
End Sub

Private Sub lst_Buddy_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_Buddy_Click()
    On Error Resume Next
    txt_To = lst_Buddy.SelectedItem.Text
End Sub

Private Sub lst_Buddy_DblClick()
    On Error Resume Next
    If Not lst_Buddy.SelectedItem.Text = "" Then
        FindPm lst_Buddy.SelectedItem.Text, False, True
        Unload Me
    End If
End Sub
