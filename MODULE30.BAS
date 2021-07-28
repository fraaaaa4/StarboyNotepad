Attribute VB_Name = "Module3"
Public Sub salva_var()
Dim fileName As String
Dim salvato As Integer
salvato = 0
End Sub

Public Sub salva_bxy()
Form1.CommonDialog1.ShowSave
fileName = Form1.CommonDialog1.fileName + ".doc"
Open fileName For Append As #1
          Print #1, Form1.Text1.Text, Form1.Text1.BackColor, Form1.Text1.Font.Size, Form1.Text1.ForeColor
          Close #1
          salvato = 1
End Sub

Public Sub salva_axa()

 If (salvato = 0) Then
                   Form1.CommonDialog1.ShowSave
                    fileName = Form1.CommonDialog1.fileName + ".doc"
                    Open fileName For Append As #1
                    Print #1, Form1.Text1.Text, Form1.Text1.BackColor
                    Close #1
          ElseIf (salvato = 1) Then
                    Open fileName For Append As #1
                    Print #1, Form1.Text1.Text
                    Close #1
          Else
          MsgBox ("Error, there's nothing to save!")
          
         End If
End Sub

