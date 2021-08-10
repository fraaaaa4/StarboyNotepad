Attribute VB_Name = "Module3"
Public Sub salva_var()
Dim filename As String
Dim salvato As Integer
salvato = 0
End Sub

Public Sub salva_bxy()
Form1.CommonDialog1.InitDir = "%USERPROFILE%"
Form1.CommonDialog1.Filter = "DOC File (*.doc)|*.doc|TXT File (*.txt)|*.txt|LOG file (*.log)|*.log|INI File (*.ini)|*.ini|DLL Files (*.dll)|*.dll|All Files (*.*)|*.*"
Form1.CommonDialog1.DialogTitle = "Save File"
Form1.CommonDialog1.ShowSave
filename = Form1.CommonDialog1.filename + ".doc"
Open filename For Append As #1
          Print #1, Form1.RichTextBox1.Text, Form1.RichTextBox1.BackColor, Form1.RichTextBox1.Font.Size, Form1.RichTextBox1.SelColor
          Close #1
          salvato = 1
End Sub

Public Sub salva_axa()

 If (salvato = 0) Then
 Form1.CommonDialog1.InitDir = "%USERPROFILE%"
Form1.CommonDialog1.Filter = "DOC File (*.doc)|*.doc|TXT File (*.txt)|*.txt|LOG file (*.log)|*.log|INI File (*.ini)|*.ini|DLL Files (*.dll)|*.dll|All Files (*.*)|*.*"
Form1.CommonDialog1.DialogTitle = "Save File"
                   Form1.CommonDialog1.ShowSave
                    filename = Form1.CommonDialog1.filename + ".doc"
                    Open filename For Append As #1
                    Print #1, Form1.RichTextBox1.Text, Form1.RichTextBox1.BackColor
                    Close #1
          ElseIf (salvato = 1) Then
                    Open filename For Append As #1
                    Print #1, Form1.RichTextBox1.Text
                    Close #1
          Else
          MsgBox ("Error, there's nothing to save!")
          
         End If
End Sub

