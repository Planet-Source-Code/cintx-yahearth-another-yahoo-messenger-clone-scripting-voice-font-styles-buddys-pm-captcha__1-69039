VERSION 5.00
Begin VB.Form frm_Add 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add a Contact"
   ClientHeight    =   2640
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
   Icon            =   "frm_Add.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmd_Add 
      Caption         =   "Add"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txt_Message 
      Height          =   1335
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txt_From 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txt_ID 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Yahoo! ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Add_Click()
    SendData AddContact(YMSG.strUser, txt_From, "Friends", txt_Message, txt_ID)
    cmd_Add.Enabled = False
End Sub

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub
