Attribute VB_Name = "Module6"
Public Sub DefaultFontOpen()
Form1.CommonDialog1.flags = 1
Form1.CommonDialog1.ShowFont
Form1.RichTextBox1.Font = Form1.CommonDialog1.FontName
Form1.RichTextBox1.Font.Size = Form1.CommonDialog1.FontSize
Form1.RichTextBox1.Font.Bold = Form1.CommonDialog1.FontBold
Form1.RichTextBox1.Font.Italic = Form1.CommonDialog1.FontItalic
Form1.RichTextBox1.Font.Underline = Form1.CommonDialog1.FontUnderline
End Sub

