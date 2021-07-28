Attribute VB_Name = "Module2"



Public Sub AddFile()
Dim ContentFile As String
On Error GoTo A
Form1.CommonDialog1.ShowOpen
Open Form1.CommonDialog1.fileName For Input As #1
Do Until EOF(1)
Input #1, ContentFile
Form1.Text1.Text = Form1.Text1.Text + ContentFile
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
Form1.Text1.Text = ContentFile
Form1.Caption = "Starboy Notepad" + " " + "-" + " " + Form1.CommonDialog1.fileName
Form1.Label63.Caption = Form1.CommonDialog1.fileName
Loop
Close #1
A:
End Sub

Public Sub FontOpen()
Form1.CommonDialog1.flags = 1
Form1.CommonDialog1.ShowFont
Form1.Text1.Font = Form1.CommonDialog1.FontName
Form1.Text1.Font.Size = Form1.CommonDialog1.FontSize
Form1.Text1.Font.Bold = Form1.CommonDialog1.FontBold
Form1.Text1.Font.Italic = Form1.CommonDialog1.FontItalic
Form1.Text1.Font.Underline = Form1.CommonDialog1.FontUnderline
Form1.Combo1.Text = Form1.CommonDialog1.FontName
Form1.Combo2.Text = Form1.CommonDialog1.FontSize
End Sub


Public Sub fuck()
Form1.CommonDialog1.flags = 1
Form1.CommonDialog1.ShowFont
Form1.Text1.Font.Name = Form1.CommonDialog1.FontName
Form1.Text1.Font.Size = Form1.CommonDialog1.FontSize
Form1.Text1.Font.Bold = Form1.CommonDialog1.FontBold
Form1.Text1.Font.Italic = Form1.CommonDialog1.FontItalic
Form1.Text1.Font.Underline = Form1.CommonDialog1.FontUnderline

End Sub


