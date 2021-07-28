Attribute VB_Name = "Module2"



Public Sub AddFile()
Dim ContentFile As String
On Error GoTo A
Form1.CommonDialog1.ShowOpen
Open Form1.CommonDialog1.fileName For Input As #1
Do Until EOF(1)
Input #1, ContentFile
Form1.RichTextBox1.Text = Form1.RichTextBox1.Text + ContentFile
Loop
Close #1
A:
End Sub


Public Sub OpenFile()
Dim ContentFile As String
On Error GoTo A
Form1.CommonDialog1.ShowOpen
Open Form1.CommonDialog1.fileName For Input As #1
Do Until EOF(1)
Input #1, ContentFile
Form1.RichTextBox1.SelText = ContentFile
Loop
Close #1
A:
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
frmOptions.RichTextBox1.Font = Form1.CommonDialog1.FontName
frmOptions.RichTextBox1.Font.Size = Form1.CommonDialog1.FontSize
frmOptions.RichTextBox1.Font.Bold = Form1.CommonDialog1.FontBold
frmOptions.RichTextBox1.Font.Italic = Form1.CommonDialog1.FontItalic
frmOptions.RichTextBox1.Font.Underline = Form1.CommonDialog1.FontUnderline
End Sub


