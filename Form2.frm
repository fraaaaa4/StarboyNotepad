VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPUForm 
      Caption         =   "Form Popup"
      Visible         =   0   'False
      Begin VB.Menu undo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu redo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu fuck 
         Caption         =   "-"
      End
      Begin VB.Menu info 
         Caption         =   "File Info"
      End
      Begin VB.Menu spell 
         Caption         =   "Spell Check"
      End
      Begin VB.Menu find 
         Caption         =   "Find"
      End
      Begin VB.Menu tasks 
         Caption         =   "Tasks"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub copy_Click()
Clipboard.Clear
Clipboard.SetText Form1.RichTextBox1.SelText, vbCFRTF
Form1.RichTextBox2.Text = ""
Form1.RichTextBox2.TextRTF = Form1.RichTextBox1.SelRTF
End Sub

Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetText Form1.RichTextBox1.SelText, vbCFRTF
Form1.RichTextBox2.Text = ""
Form1.RichTextBox2.TextRTF = Form1.RichTextBox1.SelRTF
Form1.RichTextBox1.SelRTF = ""
End Sub

Private Sub find_Click()
SendKeys "^A"
frmFind.Show
End Sub

Private Sub info_Click()
Form1.Frame8.Visible = True
End Sub

Private Sub paste_Click()
 Form1.RichTextBox1.SelRTF = Form1.RichTextBox2.TextRTF
End Sub

Private Sub redo_Click()
SendKeys "^Y"
End Sub

Private Sub spell_Click()
Spellcheck
End Sub

Private Sub tasks_Click()
Form1.Frame9.Visible = True
End Sub

Private Sub undo_Click()
SendKeys "^Z"
End Sub
