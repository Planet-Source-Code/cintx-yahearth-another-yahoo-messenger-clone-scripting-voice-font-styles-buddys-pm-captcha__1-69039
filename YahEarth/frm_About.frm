VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.PictureBox pcScroll 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3735
         ScaleWidth      =   4335
         TabIndex        =   2
         Top             =   2640
         Width           =   4335
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "SASiO"
            Height          =   255
            Left            =   840
            TabIndex        =   17
            Top             =   3480
            Width           =   3495
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "PHPFreak"
            Height          =   255
            Left            =   840
            TabIndex        =   16
            Top             =   3240
            Width           =   3495
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "BliZzArD / OloX"
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   3000
            Width           =   3495
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Hyochan - Great Music"
            Height          =   255
            Left            =   840
            TabIndex        =   14
            Top             =   2760
            Width           =   3495
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "MaddoxX"
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   2520
            Width           =   3495
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Bolle"
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   2280
            Width           =   3495
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Syndrom"
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   2040
            Width           =   3495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Burlex"
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Author: cIntX / CiPH"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   4455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Date: 11-09-2006"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Credits: "
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Matthew Robertson"
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "German-Boy"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Yazak"
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Gareth"
            Height          =   255
            Left            =   840
            TabIndex        =   3
            Top             =   1560
            Width           =   3495
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YahEarth v1.2 Alpha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
    pcScroll.Top = pcScroll.Top - 5
    If pcScroll.Top = -pcScroll.Height Then
        pcScroll.Top = (pcScroll.Height + 500)
    End If
End Sub
