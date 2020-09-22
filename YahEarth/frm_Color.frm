VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_Color 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Color.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Convert 
      Caption         =   ">"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   1830
      Width           =   255
   End
   Begin VB.TextBox txt_Hex 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "RGB"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin ComctlLib.Slider slider_R 
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Max             =   255
         TickFrequency   =   15
      End
      Begin ComctlLib.Slider slider_G 
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Max             =   255
         TickFrequency   =   15
      End
      Begin ComctlLib.Slider slider_B 
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Max             =   255
         TickFrequency   =   15
      End
      Begin VB.Label lbl_Preview 
         BackColor       =   &H00000000&
         Height          =   975
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "B:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "G:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "R:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Frm As Form
Dim Box As TextBox
Dim FromBox As Boolean

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_Convert_Click()
    txt_Hex = Trim(txt_Hex)
    Call Hex2RGB_2(txt_Hex)
End Sub

Private Sub cmd_OK_Click()
    If FromBox = True Then
        Box.Text = RGB2Hex(RGB(slider_R.Value, slider_G.Value, slider_B.Value))
    Else
        Frm.txt_Message.SelColor = RGB(slider_R.Value, slider_G.Value, slider_B.Value)
    End If
    Unload Me
End Sub

Private Sub slider_R_Change()
    lbl_Preview.BackColor = RGB(slider_R.Value, slider_G.Value, slider_B.Value)
    txt_Hex = RGB2Hex(RGB(slider_R.Value, slider_G.Value, slider_B.Value))
End Sub

Private Sub slider_G_Change()
    lbl_Preview.BackColor = RGB(slider_R.Value, slider_G.Value, slider_B.Value)
    txt_Hex = RGB2Hex(RGB(slider_R.Value, slider_G.Value, slider_B.Value))
End Sub

Private Sub slider_B_Change()
    lbl_Preview.BackColor = RGB(slider_R.Value, slider_G.Value, slider_B.Value)
    txt_Hex = RGB2Hex(RGB(slider_R.Value, slider_G.Value, slider_B.Value))
End Sub

Sub LoadColorForm(frmForm As Form, Optional UseBox As Boolean, Optional colorBox As TextBox)
    Set Frm = frmForm
    Me.Show
    If UseBox = True Then
        Set Box = colorBox
        FromBox = True
    End If
End Sub

Sub Hex2RGB_2(ByVal strHex As String)
    Dim R As String, G As String, B As String
    If Left(strHex, 1) = "#" Then strHex = Mid(strHex, 2)
    R = Mid(strHex, 1, 2)
    G = Mid(strHex, 3, 2)
    B = Mid(strHex, 5, 2)
    If Len(R) = 1 Then R = "0" & R
    If Len(G) = 1 Then G = "0" & G
    If Len(B) = 1 Then B = "0" & B
    R = Val("&h" & R)
    G = Val("&h" & G)
    B = Val("&h" & B)
    slider_R.Value = R
    slider_G.Value = G
    slider_B.Value = B
    lbl_Preview.BackColor = RGB(R, G, B)
End Sub

