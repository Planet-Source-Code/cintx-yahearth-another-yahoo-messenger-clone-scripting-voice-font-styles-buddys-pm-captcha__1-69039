VERSION 5.00
Begin VB.Form frm_Smileys 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_Smiley2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lbl_Smiley 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   2280
      Width           =   495
   End
   Begin VB.Shape mOver 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   86
      Left            =   3240
      Picture         =   "frm_Smileys.frx":0000
      Top             =   2280
      Width           =   780
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   85
      Left            =   2640
      Picture         =   "frm_Smileys.frx":20DE
      Top             =   2280
      Width           =   600
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   84
      Left            =   2280
      Picture         =   "frm_Smileys.frx":3A17
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   83
      Left            =   1680
      Picture         =   "frm_Smileys.frx":3F61
      Top             =   2280
      Width           =   450
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   82
      Left            =   1320
      Picture         =   "frm_Smileys.frx":5FC7
      Top             =   2280
      Width           =   420
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   81
      Left            =   600
      Picture         =   "frm_Smileys.frx":6923
      Top             =   2280
      Width           =   660
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   80
      Left            =   120
      Picture         =   "frm_Smileys.frx":8153
      Top             =   2280
      Width           =   420
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   79
      Left            =   4320
      Picture         =   "frm_Smileys.frx":8F7A
      Top             =   1920
      Width           =   465
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   78
      Left            =   3840
      Picture         =   "frm_Smileys.frx":9D9C
      Top             =   1920
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   77
      Left            =   3240
      Picture         =   "frm_Smileys.frx":A17D
      Top             =   1920
      Width           =   390
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   76
      Left            =   2640
      Picture         =   "frm_Smileys.frx":B27E
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   75
      Left            =   2040
      Picture         =   "frm_Smileys.frx":BDBD
      Top             =   1920
      Width           =   540
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   74
      Left            =   1680
      Picture         =   "frm_Smileys.frx":D608
      Top             =   1920
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   73
      Left            =   1320
      Picture         =   "frm_Smileys.frx":D7B2
      Top             =   1920
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   72
      Left            =   840
      Picture         =   "frm_Smileys.frx":D97F
      Top             =   1920
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   71
      Left            =   480
      Picture         =   "frm_Smileys.frx":DB62
      Top             =   1920
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   70
      Left            =   120
      Picture         =   "frm_Smileys.frx":DD07
      Top             =   1920
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   69
      Left            =   4920
      Picture         =   "frm_Smileys.frx":EC31
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   68
      Left            =   4920
      Picture         =   "frm_Smileys.frx":F453
      Top             =   1560
      Width           =   390
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   67
      Left            =   4560
      Picture         =   "frm_Smileys.frx":101D7
      Top             =   1560
      Width           =   330
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   66
      Left            =   4200
      Picture         =   "frm_Smileys.frx":10BCD
      Top             =   1560
      Width           =   330
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   65
      Left            =   3840
      Picture         =   "frm_Smileys.frx":11646
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   64
      Left            =   3480
      Picture         =   "frm_Smileys.frx":12D33
      Top             =   1560
      Width           =   330
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   63
      Left            =   3120
      Picture         =   "frm_Smileys.frx":13901
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   62
      Left            =   2760
      Picture         =   "frm_Smileys.frx":13D27
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   61
      Left            =   2400
      Picture         =   "frm_Smileys.frx":145BA
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   60
      Left            =   2040
      Picture         =   "frm_Smileys.frx":1475C
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   59
      Left            =   1680
      Picture         =   "frm_Smileys.frx":14E0B
      Top             =   1560
      Width           =   300
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   58
      Left            =   1200
      Picture         =   "frm_Smileys.frx":15567
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   57
      Left            =   720
      Picture         =   "frm_Smileys.frx":16001
      Top             =   1560
      Width           =   450
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   56
      Left            =   480
      Picture         =   "frm_Smileys.frx":16717
      Top             =   1560
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   55
      Left            =   120
      Picture         =   "frm_Smileys.frx":16B8C
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   54
      Left            =   4920
      Picture         =   "frm_Smileys.frx":17158
      Top             =   1200
      Width           =   375
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   53
      Left            =   4560
      Picture         =   "frm_Smileys.frx":178C7
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   52
      Left            =   4200
      Picture         =   "frm_Smileys.frx":17A4C
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   51
      Left            =   3840
      Picture         =   "frm_Smileys.frx":17B5F
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   50
      Left            =   3480
      Picture         =   "frm_Smileys.frx":17E99
      Top             =   1200
      Width           =   315
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   49
      Left            =   3120
      Picture         =   "frm_Smileys.frx":18350
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   48
      Left            =   2760
      Picture         =   "frm_Smileys.frx":18AD4
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   47
      Left            =   2400
      Picture         =   "frm_Smileys.frx":19064
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   46
      Left            =   2040
      Picture         =   "frm_Smileys.frx":19797
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   45
      Left            =   1680
      Picture         =   "frm_Smileys.frx":19AB0
      Top             =   1200
      Width           =   360
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   44
      Left            =   1200
      Picture         =   "frm_Smileys.frx":1A0A7
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   43
      Left            =   840
      Picture         =   "frm_Smileys.frx":1A724
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   42
      Left            =   480
      Picture         =   "frm_Smileys.frx":1AAFF
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   41
      Left            =   0
      Picture         =   "frm_Smileys.frx":1AD70
      Top             =   1200
      Width           =   540
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   40
      Left            =   4920
      Picture         =   "frm_Smileys.frx":1B5F6
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   39
      Left            =   4560
      Picture         =   "frm_Smileys.frx":1BB09
      Top             =   840
      Width           =   360
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   38
      Left            =   4200
      Picture         =   "frm_Smileys.frx":1C044
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   37
      Left            =   3840
      Picture         =   "frm_Smileys.frx":1C42B
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   36
      Left            =   3480
      Picture         =   "frm_Smileys.frx":1C68A
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   35
      Left            =   3120
      Picture         =   "frm_Smileys.frx":1CDA9
      Top             =   840
      Width           =   570
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   34
      Left            =   2760
      Picture         =   "frm_Smileys.frx":1DD11
      Top             =   840
      Width           =   360
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   33
      Left            =   2400
      Picture         =   "frm_Smileys.frx":1E735
      Top             =   840
      Width           =   420
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   32
      Left            =   2040
      Picture         =   "frm_Smileys.frx":1F4AC
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   31
      Left            =   1680
      Picture         =   "frm_Smileys.frx":1F8AE
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   30
      Left            =   1320
      Picture         =   "frm_Smileys.frx":20353
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   29
      Left            =   840
      Picture         =   "frm_Smileys.frx":20A7A
      Top             =   840
      Width           =   360
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   28
      Left            =   480
      Picture         =   "frm_Smileys.frx":21139
      Top             =   840
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   27
      Left            =   120
      Picture         =   "frm_Smileys.frx":21516
      Top             =   840
      Width           =   315
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   26
      Left            =   4920
      Picture         =   "frm_Smileys.frx":21A52
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   25
      Left            =   4560
      Picture         =   "frm_Smileys.frx":21F1A
      Top             =   480
      Width           =   360
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   24
      Left            =   4080
      Picture         =   "frm_Smileys.frx":226DE
      Top             =   480
      Width           =   450
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   23
      Left            =   3540
      Picture         =   "frm_Smileys.frx":24571
      Top             =   480
      Width           =   450
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   22
      Left            =   3120
      Picture         =   "frm_Smileys.frx":271DD
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   21
      Left            =   2760
      Picture         =   "frm_Smileys.frx":2751E
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   20
      Left            =   2400
      Picture         =   "frm_Smileys.frx":2778F
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   19
      Left            =   1920
      Picture         =   "frm_Smileys.frx":27A21
      Top             =   480
      Width           =   330
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   18
      Left            =   1560
      Picture         =   "frm_Smileys.frx":2832D
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   17
      Left            =   1200
      Picture         =   "frm_Smileys.frx":292DE
      Top             =   480
      Width           =   510
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   16
      Left            =   840
      Picture         =   "frm_Smileys.frx":2A396
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   15
      Left            =   480
      Picture         =   "frm_Smileys.frx":2A855
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   14
      Left            =   120
      Picture         =   "frm_Smileys.frx":2AC5A
      Top             =   480
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   13
      Left            =   4800
      Picture         =   "frm_Smileys.frx":2AF8F
      Top             =   120
      Width           =   510
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   12
      Left            =   4560
      Picture         =   "frm_Smileys.frx":2C23D
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   11
      Left            =   4200
      Picture         =   "frm_Smileys.frx":2C8CD
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   10
      Left            =   3840
      Picture         =   "frm_Smileys.frx":2D1E7
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   9
      Left            =   3480
      Picture         =   "frm_Smileys.frx":2D718
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   8
      Left            =   3120
      Picture         =   "frm_Smileys.frx":2DA71
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   7
      Left            =   2760
      Picture         =   "frm_Smileys.frx":2E0E6
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   6
      Left            =   2400
      Picture         =   "frm_Smileys.frx":2EA05
      Top             =   120
      Width           =   300
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   5
      Left            =   1800
      Picture         =   "frm_Smileys.frx":2F4B9
      Top             =   120
      Width           =   630
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   4
      Left            =   1560
      Picture         =   "frm_Smileys.frx":30265
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   3
      Left            =   1200
      Picture         =   "frm_Smileys.frx":306EB
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   2
      Left            =   840
      Picture         =   "frm_Smileys.frx":3090F
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   1
      Left            =   480
      Picture         =   "frm_Smileys.frx":30D04
      Top             =   120
      Width           =   270
   End
   Begin VB.Image img_Smiley 
      Height          =   270
      Index           =   0
      Left            =   120
      Picture         =   "frm_Smileys.frx":310F9
      Top             =   120
      Width           =   270
   End
End
Attribute VB_Name = "frm_Smileys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Frm As Form

Private Sub Form_Load()
    ExecuteScript 7, , , , Me.Name
    InitCommonControls
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub img_Smiley_Click(Index As Integer)
    Dim strList() As String
    Dim strSmiley() As String
    Dim strSmileys As String
    Dim X As Integer
    
    'On Error Resume Next
    
    strSmileys = LoadTextFile(App.Path & "\Resources\Pictures\Smileys\Smileys.txt")
    strList = Split(strSmileys, vbCrLf)
    For X = 0 To UBound(strList) - 1
        strSmiley = Split(strList(X), " ")
        If LCase(strSmiley(0)) = "smiley" & (Index + 1) Then
            'Process Send to Form
            Frm.txt_Message.SelText = strSmiley(1)
            Frm.SetFocus
            Unload Me
            Exit For
        End If
    Next X
End Sub

Private Sub img_Smiley_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim strList() As String
    Dim strSmiley() As String
    Dim strSmileys As String
    Dim i As Integer
    
    On Error Resume Next
    
    strSmileys = LoadTextFile(App.Path & "\Resources\Pictures\Smileys\Smileys.txt")
    strList = Split(strSmileys, vbCrLf)
    For i = 0 To UBound(strList) - 1
        strSmiley = Split(strList(i), " ")
        If LCase(strSmiley(0)) = "smiley" & (Index + 1) Then
            'Process Send to Form
            If UBound(strSmiley) > 1 Then
                lbl_Smiley.Caption = strSmiley(1)
                lbl_Smiley2.Caption = strSmiley(2)
            Else
                lbl_Smiley.Caption = strSmiley(1)
                lbl_Smiley2.Caption = ""
            End If
        End If
    Next i
    
    If Index = 41 Then
        mOver.Left = img_Smiley(27).Left
        mOver.Width = img_Smiley(27).Width
    ElseIf Index = 5 Then
        mOver.Left = img_Smiley(Index).Left + 180
        mOver.Width = img_Smiley(Index).Width - 350
    ElseIf Index = 13 Then
        mOver.Left = img_Smiley(Index).Left + 100
        mOver.Width = img_Smiley(Index).Width - 200
    ElseIf Index = 17 Then
        mOver.Left = img_Smiley(Index).Left
        mOver.Width = img_Smiley(Index).Width - 200
    ElseIf Index = 35 Then
        mOver.Left = img_Smiley(Index).Left
        mOver.Width = img_Smiley(Index).Width - 240
    Else
        mOver.Left = img_Smiley(Index).Left
        mOver.Width = img_Smiley(Index).Width
    End If
    mOver.Top = img_Smiley(Index).Top
    mOver.Height = img_Smiley(Index).Height
End Sub

Sub OpenSmileys(FrmFrom As Form)
    Me.Show
    Set Frm = FrmFrom
End Sub
