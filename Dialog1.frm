VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Dialog1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Clipboard Text - Starboy Notepad"
   ClientHeight    =   3195
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      MousePointer    =   3
      Appearance      =   0
      TextRTF         =   $"Dialog1.frx":0000
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_GotFocus()
RichTextBox2.TextRTF = Form1.RichTextBox2.TextRTF
End Sub

Private Sub Form_Load()
Dialog1.Top = Form1.Top + 40
Dialog1.Left = Form1.Left + 40
RichTextBox2.TextRTF = Form1.RichTextBox2.TextRTF
End Sub

Private Sub Form_Resize()
RichTextBox2.Width = Dialog1.Width - 15
RichTextBox2.Height = Dialog1.Height - 15
End Sub
