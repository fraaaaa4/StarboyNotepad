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

Public Sub Spellcheck()

    Dim objWord As Object
    Dim objDoc  As Object
    Dim strResult As String
    Const QUOTE = """"
    
    On Error GoTo ErrorRoutine

    App.OleRequestPendingTimeout = 999999
    Set objWord = GetObject("Word.Application")
    If TypeName(objWord) <> "Nothing" Then
        ' Word is already open
        Set objWord = GetObject(, "Word.Application")
    Else
        ' Create an instance of Word
        Set objWord = CreateObject("Word.Application")
    End If

    Select Case objWord.version
        'Office 2000 and later
        Case "9.0", "10.0", "11.0", "14.0", "15.0"
            Set objDoc = objWord.Documents.Add(, , 1, True)
        'Office 97
        Case "8.0"
            Set objDoc = objWord.Documents.Add
        Case Else
            MsgBox "Sorry but your version of Word seems to be " & QUOTE & objWord.version _
                   & QUOTE & " and that version is not currently supported.", vbOKOnly + vbExclamation, "Spelling Checker"
            Exit Sub
    End Select

    objDoc.Content = Form1.RichTextBox1.Text
    objDoc.CheckSpelling
    objWord.Visible = False

    strResult = Left(objDoc.Content, Len(objDoc.Content) - 1)
    ' Reformat carriage returns
    strResult = Replace(strResult, Chr(13), Chr(13) & Chr(10))
    
    If Form1.RichTextBox1.Text = strResult Then
        ' There were no errors, so give the user a
        ' visual signal that something happened
        MsgBox "No changes made", vbInformation + vbOKOnly, "Spelling Checker"
    End If
    
    'Clean up
    objDoc.Close False
    Set objDoc = Nothing
    objWord.Application.Quit True
    Set objWord = Nothing

    ' Replace the selected text with the corrected text. It's important that
    ' this be done after the "Clean Up" because otherwise there are problems
    ' with the screen not repainting
    Form1.RichTextBox1.TextRTF = strResult

    Exit Sub


ErrorRoutine:

    Select Case Err.Number
        Case -2147221020
            ' There's no instance of Word so continue processing in order to create one
            Resume Next
        Case 429
            MsgBox "Word must be installed in order for this code to work", vbCritical + vbOKOnly, "Spelling Checker"
    End Select

Form1.Image77.Visible = True
End Sub
