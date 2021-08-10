Attribute VB_Name = "Module7"
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lparam As Any) As Long
Public Const EM_GETLINECOUNT As Long = &HBA
Public Sub getemboys()
Dim linecount As Integer
Dim count As String
Dim i As Integer
  linecount = SendMessage(Form1.RichTextBox1.hWnd, _
                  EM_GETLINECOUNT, 0, 0)
                  Form1.List1.Clear
                  For i = "1" To linecount
                  Form1.List1.AddItem (i)
Form1.Label63.Caption = linecount
Next
End Sub


Public Sub getemfuckingboys()
Dim linecount As Integer
  linecount = SendMessage(Form1.RichTextBox1.hWnd, _
                  EM_GETLINECOUNT, 0, 0)
Form1.Label63.Caption = linecount
End Sub


Public Sub Notoptionalmyass(linecount As Integer)
    Dim i As Integer

    Form1.List1.Clear
    For i = "1" To linecount
        Form1.List1.AddItem ("i")
    Next
End Sub


