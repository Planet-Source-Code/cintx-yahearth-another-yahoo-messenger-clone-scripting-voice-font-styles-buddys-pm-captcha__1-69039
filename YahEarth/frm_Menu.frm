VERSION 5.00
Begin VB.Form frm_Menu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   Icon            =   "frm_Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu menu_List 
      Caption         =   "List"
      Begin VB.Menu mnu_Ignore 
         Caption         =   "Ignore User"
      End
      Begin VB.Menu mnu_Mute 
         Caption         =   "Mute User"
      End
      Begin VB.Menu mnu_1_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_PM 
         Caption         =   "Send PM"
      End
   End
   Begin VB.Menu mnu_Status 
      Caption         =   "Status"
      Begin VB.Menu mnu_Online 
         Caption         =   "Online"
      End
      Begin VB.Menu mnu_2_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Busy 
         Caption         =   "Busy"
      End
      Begin VB.Menu mnu_SteppedOut 
         Caption         =   "Stepped Out"
      End
      Begin VB.Menu mnu_BRB 
         Caption         =   "Be Right Back"
      End
      Begin VB.Menu mnu_notatdesk 
         Caption         =   "Not at My Desk"
      End
      Begin VB.Menu mnu_onPhone 
         Caption         =   "On The Phone"
      End
      Begin VB.Menu mnu_2_line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_invisible 
         Caption         =   "Invisible"
      End
   End
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_BRB_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        SendData Status_BrB
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: Be Right Back"
    End If
End Sub

Private Sub mnu_Busy_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        SendData Status_Busy
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: Busy"
    End If
End Sub

Private Sub mnu_Ignore_Click()
    If mnu_Ignore.Caption = "Ignore User" Then
        AddIgnore frm_Main.lst_User.SelectedItem.Text
    Else
        RemoveIgnore frm_Main.lst_User.SelectedItem.Text
    End If
End Sub

Private Sub mnu_invisible_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        'SendData Status_Invisible
        DoEvents
        SendData Status_Invisible2(YMSG.strUser)
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: Invisible"
    End If
End Sub

Private Sub mnu_Mute_Click()
    If mnu_Mute.Caption = "Mute User" Then
        frm_Main.Voice.muteSource frm_Main.lst_User.SelectedItem.Tag, frm_Main.lst_User.SelectedItem.Text
        frm_Main.lst_User.SelectedItem.SmallIcon = 3
    Else
        frm_Main.lst_User.SelectedItem.SmallIcon = 2
        frm_Main.Voice.unmuteSource frm_Main.lst_User.SelectedItem.Tag, frm_Main.lst_User.SelectedItem.Text
    End If
End Sub

Private Sub mnu_notatdesk_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        SendData Status_NotAtDesk
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: Not at My Desk"
    End If
End Sub

Private Sub mnu_Online_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        SendData Status_Online
        DoEvents
        SendData Status_Available
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: Online"
    End If
End Sub

Private Sub mnu_onPhone_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        SendData Status_OnPhone
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: On the Phone"
    End If
End Sub

Private Sub mnu_PM_Click()
    FindPm frm_Main.lst_User.SelectedItem.Text, , True
End Sub

Private Sub mnu_SteppedOut_Click()
    If YMSG.strUser = "" Then
        frm_Login.Show
    Else
        SendData Status_SteppedOut
        frm_Main.StatusBar1.Panels(2).Text = "Online Status: Stepped Out"
    End If
End Sub
