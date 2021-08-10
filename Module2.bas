Attribute VB_Name = "Module2"



Public Sub AddFile()
Dim ContentFile As String
On Error GoTo a
Form1.CommonDialog1.InitDir = "%USERPROFILE%"
Form1.CommonDialog1.Filter = "DOC File (*.doc)|*.doc|TXT File (*.txt)|*.txt|LOG file (*.log)|*.log|INI File (*.ini)|*.ini|DLL Files (*.dll)|*.dll|All Files (*.*)|*.*"
Form1.CommonDialog1.DialogTitle = "Add file"
Form1.CommonDialog1.ShowOpen
Open Form1.CommonDialog1.filename For Input As #1


Do Until EOF(1)
Input #1, ContentFile
Form1.RichTextBox1.Text = Form1.RichTextBox1.Text + ContentFile
Loop
Close #1
a:
End Sub


Public Sub OpenFile()
Dim ContentFile As String


On Error GoTo a
Form1.CommonDialog1.InitDir = "%USERPROFILE%"
Form1.CommonDialog1.Filter = "DOC File (*.doc)|*.doc|TXT File (*.txt)|*.txt|LOG file (*.log)|*.log|INI File (*.ini)|*.ini|DLL Files (*.dll)|*.dll|All Files (*.*)|*.*"
Form1.CommonDialog1.DialogTitle = "Open File"
Form1.CommonDialog1.ShowOpen
Open Form1.CommonDialog1.filename For Input As #1
Do Until EOF(1)
Input #1, ContentFile
Form1.RichTextBox1.TextRTF = ContentFile
Form1.Caption = "Starboy Notepad - " + Form1.CommonDialog1.FileTitle
Form1.Label2.Caption = "Starboy Notepad - " + Form1.CommonDialog1.FileTitle
Form1.Label79.Caption = Form1.CommonDialog1.FileTitle
Form1.Label81.Caption = Form1.CommonDialog1.filename
Loop
Close #1
a:
End Sub

Public Sub FontOpen()
Form1.CommonDialog1.flags = 1
Form1.CommonDialog1.ShowFont
Form1.RichTextBox1.SelFontName = Form1.CommonDialog1.FontName
Form1.RichTextBox1.SelFontSize = Form1.CommonDialog1.FontSize
Form1.RichTextBox1.SelBold = Form1.CommonDialog1.FontBold
Form1.RichTextBox1.SelItalic = Form1.CommonDialog1.FontItalic
Form1.RichTextBox1.SelUnderline = Form1.CommonDialog1.FontUnderline
End Sub


Public Sub fuck()
Form1.CommonDialog1.flags = 1
Form1.CommonDialog1.ShowFont
Form1.RichTextBox1.SelFontName = Form1.CommonDialog1.FontName
Form1.RichTextBox1.SelFontSize = Form1.CommonDialog1.FontSize
Form1.RichTextBox1.SelBold = Form1.CommonDialog1.FontBold
Form1.RichTextBox1.SelItalic = Form1.CommonDialog1.FontItalic
Form1.RichTextBox1.SelUnderline = Form1.CommonDialog1.FontUnderline

End Sub


