VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   1005
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsMouseDown As Boolean
Dim MouseDownX As Integer
Dim MouseDownY As Integer

Private Sub Form_Click()
    Me.BackColor = RGB(250, 0, 0)
    Timer1.Enabled = False
End Sub

Private Sub Form_DblClick()
    Me.BackColor = RGB(200, 200, 0)
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    ' Restore Form Position
    FormLoadPosition Me

    ' Transparent at ratio 140/255
    Me.BackColor = RGB(250, 0, 0)
    ActiveTransparency Me, True, False, 100, Me.BackColor
    SetTopMostWindow Me.hWnd, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        IsMouseDown = True
        MouseDownX = x
        MouseDownY = y
        Debug.Print "Form_MouseDown"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "Form_MouseMove", x, y
    If IsMouseDown Then
        Me.Move Me.Left - MouseDownX + x, Me.Top - MouseDownY + y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' Right Click
        Timer1.Enabled = False
        PopupMenu Form2.mnuFile
    ElseIf Button = 1 Then
        Debug.Print "Form_MouseUp"
        IsMouseDown = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormSavePosition Me
End Sub

Private Sub Timer1_Timer()
    Debug.Print "Timer1_Timer"
    SendKeys "%r"
End Sub
