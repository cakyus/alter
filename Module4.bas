Attribute VB_Name = "Module4"

Public Sub FormSavePosition(Form1 As Form)
    SaveSetting App.ProductName, Form1.Name, "Left", Form1.Left
    SaveSetting App.ProductName, Form1.Name, "Top", Form1.Top
    Debug.Print "FormSavePosition", Form1.Left, Form1.Top
End Sub

Public Sub FormLoadPosition(Form1 As Form)

    Dim iTop As Single, iLeft As Single
    
    iLeft = GetSetting(App.ProductName, Form1.Name, "Left", Form1.Left)
    iTop = GetSetting(App.ProductName, Form1.Name, "Top", Form1.Top)
    
    Debug.Print "FormLoadPosition", iLeft, iTop
    
    If iLeft = 0 And iTop = 0 Then
        ' Initial position, move top center of screen
        Form1.Move Screen.Width / 2, Screen.Height / 2
        FormSavePosition Form1
    Else
        Form1.Move iLeft, iTop
    End If
End Sub
