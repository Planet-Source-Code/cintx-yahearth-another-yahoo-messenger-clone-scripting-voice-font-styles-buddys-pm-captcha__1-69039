VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_Captcha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Yahoo! Chat Captcha"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      ExtentX         =   21193
      ExtentY         =   13150
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
End
Attribute VB_Name = "frm_Captcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowCaptcha(strLink As String)
    Me.Show
    Debug.Print "show"
    WB.Navigate2 strLink
End Sub

Private Sub WB_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    If URL = "http://captcha.chat.yahoo.com/go/captchat/close?.intl=us" Then
        Unload Me
    End If
End Sub

