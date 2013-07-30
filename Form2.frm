VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFileAbout_Click()
    MsgBox App.ProductName & _
    Chr(10) & Chr(13) & App.FileDescription & _
    Chr(10) & Chr(13) & App.Comments _
    , , App.ProductName
End Sub

Private Sub mnuFileHelp_Click()
    OpenURL "https://github.com/cakyus/alter"
End Sub

Private Sub mnuFileQuit_Click()
    Unload Form1
    Unload Me
    End
End Sub
