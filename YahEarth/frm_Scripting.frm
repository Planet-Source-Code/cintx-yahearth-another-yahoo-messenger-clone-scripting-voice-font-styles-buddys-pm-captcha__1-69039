VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_Scripting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " YahEarth - Scripts"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Scripting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Custom 
      Caption         =   "Custom Functions"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd_NewWindow 
      Caption         =   "New Window"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd_onEnd 
      Caption         =   "Application End"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd_OnStart 
      Caption         =   "Application Start"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd_IncommingPM 
      Caption         =   " Incomming PM"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmd_IncommingChatText 
      Caption         =   " Incomming Chat Text"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4920
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Line: 1"
            TextSave        =   "Line: 1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Position: 1"
            TextSave        =   "Position: 1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Selected: 0"
            TextSave        =   "Selected: 0"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Size: 0 Byte"
            TextSave        =   "Size: 0 Byte"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frame_Scripts 
      Caption         =   "Event: Incomming Data (string Data)"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5760
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   14
         Top             =   3600
         Width           =   1335
         Begin VB.CommandButton cmd_Save 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4320
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   3600
         Width           =   1335
         Begin VB.CommandButton cmd_Test 
            Caption         =   "Test Script"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1335
         End
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   3135
         Left            =   120
         TabIndex        =   9
         Tag             =   "1"
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5530
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frm_Scripting.frx":57E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBBuff 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393217
         TextRTF         =   $"frm_Scripting.frx":5862
      End
   End
   Begin VB.CommandButton cmd_DataOut 
      Caption         =   "Outgoing Data"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmd_IncommingData 
      Caption         =   " Incomming Data"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_Scripting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLast As String
Private Sub cmd_Custom_Click()
    RTB.Text = Src(8).strScript
    RTB.Tag = 8
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Custom Functions"
End Sub

Private Sub cmd_DataOut_Click()
    RTB.Text = Src(2).strScript
    RTB.Tag = 2
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: Outgoing Data (string Data)"
End Sub

Private Sub cmd_IncommingChatText_Click()
    RTB.Text = Src(3).strScript
    RTB.Tag = 3
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: Incomming Chat Text (string User, string Message)"
End Sub

Private Sub cmd_IncommingData_Click()
    RTB.Text = Src(1).strScript
    RTB.Tag = 1
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: Incomming Data (string Data)"
End Sub

Private Sub cmd_IncommingPM_Click()
    RTB.Text = Src(4).strScript
    RTB.Tag = 4
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: Incomming PM (string User, string Message)"
End Sub

Private Sub cmd_NewWindow_Click()
    RTB.Text = Src(7).strScript
    RTB.Tag = 7
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: New Window (string Window)"
End Sub

Private Sub cmd_onEnd_Click()
    RTB.Text = Src(6).strScript
    RTB.Tag = 6
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: Application End"
End Sub

Private Sub cmd_OnStart_Click()
    RTB.Text = Src(5).strScript
    RTB.Tag = 5
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
    UpdateRTBInfo
    frame_Scripts.Caption = "Event: Application Start"
End Sub

Private Sub cmd_Save_Click()
    Src(RTB.Tag).strScript = RTB.Text
    SaveTextFile App.Path & "\Resources\Scripts\" & Src(RTB.Tag).strFile, RTB.Text
    strLast = RTB.Text
    cmd_Save.Enabled = False
    RTB.SetFocus
End Sub

Private Sub cmd_Test_Click()
    Dim strOld As String
    strOld = Src(RTB.Tag).strScript
    Src(RTB.Tag).strScript = RTB.Text
    If cmd_Test.Caption = "Test Script" Then
        cmd_Test.Caption = "Stop Script"
        ExecuteScript RTB.Tag, "Test Data", "Test User", "Test Message", "Test Window"
        cmd_Test.Caption = "Test Script"
    Else
        cmd_Test.Caption = "Test Script"
        frm_Main.Script.Reset
    End If
    Src(RTB.Tag).strScript = strOld
End Sub

Private Sub Form_Load()
    ExecuteScript 7, , , , Me.Name
    RTB.Text = Src(1).strScript
    strLast = RTB.Text
    StatusBar1.Panels(1).Width = Me.ScaleWidth / 4
    StatusBar1.Panels(2).Width = Me.ScaleWidth / 4
    StatusBar1.Panels(3).Width = Me.ScaleWidth / 4
    StatusBar1.Panels(4).Width = Me.ScaleWidth / 4
    UpdateRTBInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strBox As String
    If Not strLast = RTB.Text Then
        strBox = MsgBox("Save before Exit?", vbYesNo, "Save")
        Select Case strBox
            Case vbYes
                Src(RTB.Tag).strScript = RTB.Text
                SaveTextFile App.Path & "\Resources\Scripts\" & Src(RTB.Tag).strFile, RTB.Text
                strLast = RTB.Text
                cmd_Save.Enabled = False
                RTB.SetFocus
        End Select
    End If
End Sub

Private Sub RTB_Change()
    If Not RTB.Text = strLast Then
        cmd_Save.Enabled = True
    ElseIf RTB.Text = strLast Then
        cmd_Save.Enabled = False
    End If
End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        KeyCode = 0
        RTB.SelText = vbTab
    End If
End Sub

Private Sub RTB_SelChange()
    UpdateRTBInfo
End Sub

Sub UpdateRTBInfo()
    On Error Resume Next
    StatusBar1.Panels(2).Text = "Position: " & RTB.SelStart
    StatusBar1.Panels(1).Text = "Line: " & GetLineNum(RTB)
    StatusBar1.Panels(3).Text = "Selected: " & RTB.SelLength
    StatusBar1.Panels(4).Text = "Size: " & FileCalc(Len(RTB.Text))
End Sub
