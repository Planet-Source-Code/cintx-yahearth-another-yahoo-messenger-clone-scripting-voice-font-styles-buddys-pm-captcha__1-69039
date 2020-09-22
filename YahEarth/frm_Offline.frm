VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_Offline 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Offline Messages"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Offline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lst_Offline 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Message"
         Object.Width           =   2540
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   2355
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
End
Attribute VB_Name = "frm_Offline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Visible = False
    WB.Navigate "about:blank"
End Sub

Private Sub lst_Offline_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lst_Offline_Click()
    Dim blOption As Boolean
    
    On Error Resume Next
    If Not lst_Offline.SelectedItem.Text = "" Then
        WB.Navigate "about:blank"
        ProcessHTML lst_Offline.SelectedItem.Text, lst_Offline.SelectedItem.SubItems(2), WB, True
    End If
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
    Me.Visible = True
    WB.Document.write "<html>" & vbCrLf & frm_Main.txt_NoClick & vbCrLf
End Sub
