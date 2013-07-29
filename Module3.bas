Attribute VB_Name = "Module3"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
Public Sub OpenURL(URL As String)
    ShellExecute hWnd, "open", URL, vbNullString, vbNullString, Empty
End Sub

