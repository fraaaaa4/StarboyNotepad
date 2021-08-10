VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools - Starboy Notepad"
   ClientHeight    =   4050
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3000
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   4365
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2115
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1245
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   765
      Width           =   1455
   End
   Begin VB.Image Image36 
      Height          =   195
      Left            =   2040
      Picture         =   "Dialog.frx":058A
      Stretch         =   -1  'True
      ToolTipText     =   "Undo (CTRL+Z)"
      Top             =   3585
      Width           =   195
   End
   Begin VB.Image Image37 
      Height          =   195
      Left            =   2400
      Picture         =   "Dialog.frx":2F048
      Stretch         =   -1  'True
      ToolTipText     =   "Redo (CTRL+Y)"
      Top             =   3585
      Width           =   180
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2040
      X2              =   2250
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2400
      X2              =   2610
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Image Image27 
      Height          =   330
      Left            =   240
      Picture         =   "Dialog.frx":5E4FA
      Stretch         =   -1  'True
      ToolTipText     =   "Returns to the splash screen"
      Top             =   3555
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   330
      Left            =   1440
      Picture         =   "Dialog.frx":A7E44
      Stretch         =   -1  'True
      ToolTipText     =   "Settings"
      Top             =   3555
      Width           =   345
   End
   Begin VB.Image Image14 
      Height          =   195
      Left            =   240
      Picture         =   "Dialog.frx":B82AE
      Stretch         =   -1  'True
      ToolTipText     =   "Adds text to this file from another file"
      Top             =   3165
      Width           =   195
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add from file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3165
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Word count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4365
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2460
      Picture         =   "Dialog.frx":D5C60
      Stretch         =   -1  'True
      ToolTipText     =   "Changes the colour of the selected font"
      Top             =   2640
      Width           =   300
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2685
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Colour"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1380
      TabIndex        =   4
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2460
      Picture         =   "Dialog.frx":D634A
      Stretch         =   -1  'True
      ToolTipText     =   "Changes the colour of the background"
      Top             =   2160
      Width           =   300
   End
   Begin VB.Image Image8 
      Height          =   225
      Left            =   240
      Picture         =   "Dialog.frx":D6A34
      Stretch         =   -1  'True
      Top             =   2205
      Width           =   255
   End
   Begin VB.Image Image21 
      Height          =   330
      Left            =   1400
      Picture         =   "Dialog.frx":D711E
      Stretch         =   -1  'True
      ToolTipText     =   "Adds text bullets to your current row"
      Top             =   1635
      Width           =   375
   End
   Begin VB.Image Image10 
      Height          =   225
      Left            =   1920
      Picture         =   "Dialog.frx":EB9B8
      Stretch         =   -1  'True
      ToolTipText     =   "Left alignment"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image11 
      Height          =   225
      Left            =   2640
      Picture         =   "Dialog.frx":F8DE6
      Stretch         =   -1  'True
      ToolTipText     =   "Right alignment"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image12 
      Height          =   225
      Left            =   2280
      Picture         =   "Dialog.frx":106214
      Stretch         =   -1  'True
      ToolTipText     =   "Center alignment"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image15 
      Height          =   270
      Left            =   240
      Picture         =   "Dialog.frx":1134B6
      Stretch         =   -1  'True
      ToolTipText     =   "Bold"
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image Image17 
      Height          =   270
      Left            =   600
      Picture         =   "Dialog.frx":120C20
      Stretch         =   -1  'True
      ToolTipText     =   "Italic"
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image Image19 
      Height          =   270
      Left            =   960
      Picture         =   "Dialog.frx":12E852
      Stretch         =   -1  'True
      ToolTipText     =   "Underline"
      Top             =   1680
      Width           =   270
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Font size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   555
      TabIndex        =   3
      Top             =   1245
      Width           =   855
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   240
      Picture         =   "Dialog.frx":13B8D4
      Stretch         =   -1  'True
      Top             =   1245
      Width           =   255
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2115
      X2              =   2715
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   765
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   225
      Left            =   240
      Picture         =   "Dialog.frx":15A762
      Stretch         =   -1  'True
      ToolTipText     =   "Change the font of your text"
      Top             =   765
      Width           =   255
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1320
      X2              =   2760
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Image Image5 
      Height          =   195
      Left            =   240
      Picture         =   "Dialog.frx":15AE4C
      Stretch         =   -1  'True
      ToolTipText     =   "Cleans the existing file. Note: all unsaved changes will be discarded"
      Top             =   210
      Width           =   195
   End
   Begin VB.Image Image13 
      Height          =   195
      Left            =   600
      Picture         =   "Dialog.frx":15B536
      Stretch         =   -1  'True
      ToolTipText     =   "Opens a new file"
      Top             =   210
      Width           =   195
   End
   Begin VB.Image Image25 
      Height          =   195
      Left            =   960
      Picture         =   "Dialog.frx":15BC20
      Stretch         =   -1  'True
      ToolTipText     =   "Saves the current file"
      Top             =   210
      Width           =   195
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   225
      X2              =   435
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   600
      X2              =   810
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   960
      X2              =   1170
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2400
      X2              =   2610
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   2040
      X2              =   2250
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1665
      X2              =   1875
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Image Image31 
      Height          =   195
      Left            =   2400
      Picture         =   "Dialog.frx":15C30A
      Stretch         =   -1  'True
      ToolTipText     =   "Paste (CTRL+V)"
      Top             =   225
      Width           =   180
   End
   Begin VB.Image Image32 
      Height          =   195
      Left            =   2040
      Picture         =   "Dialog.frx":15C9F4
      Stretch         =   -1  'True
      ToolTipText     =   "Copy (CTRL+C)"
      Top             =   225
      Width           =   195
   End
   Begin VB.Image Image33 
      Height          =   195
      Left            =   1680
      Picture         =   "Dialog.frx":15D0DE
      Stretch         =   -1  'True
      ToolTipText     =   "Cut (CTRL+Z)"
      Top             =   225
      Width           =   195
   End
   Begin VB.Image Image34 
      Height          =   360
      Left            =   1560
      Picture         =   "Dialog.frx":15D7C8
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1245
   End
   Begin VB.Image Image30 
      Height          =   360
      Left            =   120
      Picture         =   "Dialog.frx":1C2A0A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   1200
      Picture         =   "Dialog.frx":227C4C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1650
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1995
      Picture         =   "Dialog.frx":28CE8E
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   810
   End
   Begin VB.Image Image23 
      Height          =   225
      Left            =   2640
      Picture         =   "Dialog.frx":2F20D0
      Stretch         =   -1  'True
      ToolTipText     =   "Right alignment"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image22 
      Height          =   225
      Left            =   2280
      Picture         =   "Dialog.frx":2FEF52
      Stretch         =   -1  'True
      ToolTipText     =   "Center alignment"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image24 
      Height          =   225
      Left            =   1920
      Picture         =   "Dialog.frx":30C4A0
      Stretch         =   -1  'True
      ToolTipText     =   "Left alignment"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image20 
      Height          =   270
      Left            =   960
      Picture         =   "Dialog.frx":319642
      Stretch         =   -1  'True
      ToolTipText     =   "Underline"
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image Image18 
      Height          =   270
      Left            =   600
      Picture         =   "Dialog.frx":3269FC
      Stretch         =   -1  'True
      ToolTipText     =   "Italic"
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image Image16 
      Height          =   270
      Left            =   240
      Picture         =   "Dialog.frx":333DB6
      Stretch         =   -1  'True
      ToolTipText     =   "Bold"
      Top             =   1680
      Width           =   270
   End
   Begin VB.Image Image26 
      Height          =   330
      Left            =   1400
      Picture         =   "Dialog.frx":3415A0
      Stretch         =   -1  'True
      ToolTipText     =   "Removes text bullets from your current row"
      Top             =   1630
      Width           =   375
   End
   Begin VB.Image Image28 
      Height          =   330
      Left            =   1320
      Picture         =   "Dialog.frx":355F42
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   495
   End
   Begin VB.Image Image29 
      Height          =   330
      Left            =   1440
      Picture         =   "Dialog.frx":3BB184
      Stretch         =   -1  'True
      ToolTipText     =   "Settings"
      Top             =   3555
      Width           =   345
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Image38 
      Height          =   360
      Left            =   1920
      Picture         =   "Dialog.frx":3CA78A
      Stretch         =   -1  'True
      Top             =   3525
      Width           =   765
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_GotFocus()
Text3.Text = Form1.RichTextBox1.SelFontName
Text1.Text = Form1.RichTextBox1.SelFontSize
End Sub

Private Sub Form_Load()
Dialog.Left = Form1.Left + 9800
Dialog.Top = Form1.Top
Text3.Text = Form1.RichTextBox1.SelFontName
Text1.Text = Form1.RichTextBox1.SelFontSize
If Form1.RichTextBox1.SelBold = True Then Image15.Visible = False
If Form1.RichTextBox1.SelItalic = True Then Image17.Visible = False
If Form1.RichTextBox1.SelUnderline = True Then Image19.Visible = False
If Form1.RichTextBox1.SelAlignment = 0 Then Image10.Visible = False
If Form1.RichTextBox1.SelAlignment = 1 Then Image11.Visible = False
If Form1.RichTextBox1.SelAlignment = 2 Then Image12.Visible = False
If Form1.RichTextBox1.SelBullet = 1 Then Image21.Visible = False

Dim Counts As Integer
Dim i As Integer
If Form1.RichTextBox1.SelText = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(Form1.RichTextBox1.SelText)
        If Mid(Form1.RichTextBox1.SelText, i, 1) = " " Then ' use Mid to search space
            Counts = Counts + 1
        End If
    Next
End If
Text4.Text = Counts
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line16.Visible = False
Line17.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line13.Visible = False
Line12.Visible = False
Line11.Visible = False
Line14.Visible = False
Line15.Visible = False
End Sub

Private Sub Image10_Click()
Image10.Visible = False
Image24.Visible = True
Form1.Image10.Visible = False
Form1.Image24.Visible = True
Form1.RichTextBox1.SelAlignment = 0
Image23.Visible = False
Image22.Visible = False
Image12.Visible = True
Image11.Visible = True

Form1.Image23.Visible = False
Form1.Image22.Visible = False
Form1.Image12.Visible = True
Form1.Image11.Visible = True
End Sub

Private Sub Image11_Click()
Image10.Visible = True
Image24.Visible = False
Form1.RichTextBox1.SelAlignment = 1
Image23.Visible = True
Image12.Visible = True
Image22.Visible = True
Image11.Visible = False

Form1.Image10.Visible = True
Form1.Image24.Visible = False
Form1.Image23.Visible = True
Form1.Image12.Visible = True
Form1.Image22.Visible = True
Form1.Image11.Visible = False
End Sub

Private Sub Image12_Click()
Image10.Visible = True
Image24.Visible = False
Form1.RichTextBox1.SelAlignment = 2
Image23.Visible = False
Image12.Visible = False
Image22.Visible = True
Image11.Visible = True

Form1.Image10.Visible = True
Form1.Image24.Visible = False
Form1.Image23.Visible = False
Form1.Image12.Visible = False
Form1.Image22.Visible = True
Form1.Image11.Visible = True
End Sub

Private Sub Image13_Click()
OpenFile
End Sub

Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = False
Line9.Visible = True
Line10.Visible = False
End Sub

Private Sub Image14_Click()
AddFile
End Sub

Private Sub Image15_Click()
Image15.Visible = False
Image16.Visible = True
Form1.Image15.Visible = False
Form1.Image16.Visible = True
Form1.RichTextBox1.SelBold = True
End Sub

Private Sub Image16_Click()
Image15.Visible = True
Image16.Visible = False
Form1.Image15.Visible = True
Form1.Image16.Visible = False
Form1.RichTextBox1.SelBold = False
End Sub

Private Sub Image17_Click()
Image17.Visible = False
Image18.Visible = True
Form1.Image17.Visible = False
Form1.Image18.Visible = True
Form1.RichTextBox1.SelItalic = True
End Sub

Private Sub Image18_Click()
Image17.Visible = True
Image18.Visible = False
Form1.Image17.Visible = True
Form1.Image18.Visible = False
Form1.RichTextBox1.SelItalic = False
End Sub

Private Sub Image19_Click()
Image19.Visible = False
Image20.Visible = True
Form1.Image19.Visible = False
Form1.Image20.Visible = True
Form1.RichTextBox1.SelUnderline = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line15.Visible = True
End Sub

Private Sub Image20_Click()
Image19.Visible = True
Image20.Visible = False
Form1.Image19.Visible = True
Form1.Image20.Visible = False
Form1.RichTextBox1.SelUnderline = False
End Sub

Private Sub Image21_Click()
Image21.Visible = False
Image26.Visible = True
Form1.Image21.Visible = False
Form1.Image26.Visible = True
Form1.RichTextBox1.SelBullet = 1
End Sub

Private Sub Image25_Click()
salva_bxy
End Sub

Private Sub Image25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = False
Line9.Visible = False
Line10.Visible = True
End Sub

Private Sub Image26_Click()
Image21.Visible = True
Image26.Visible = False
Form1.Image21.Visible = True
Form1.Image26.Visible = False
Form1.RichTextBox1.SelBullet = 0
End Sub

Private Sub Image27_Click()
frmSplash.Show
End Sub

Public Sub Image3_Click()
ShowColorDialogSentient
End Sub

Private Sub Image30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
End Sub

Private Sub Image31_Click()
 Form1.RichTextBox1.SelRTF = Form1.RichTextBox2.TextRTF
End Sub

Private Sub Image31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line11.Visible = True
End Sub

Private Sub Image32_Click()
Clipboard.Clear
Clipboard.SetText Form1.RichTextBox1.SelText, vbCFRTF
Form1.RichTextBox2.Text = ""
Form1.RichTextBox2.TextRTF = Form1.RichTextBox1.SelRTF
End Sub

Private Sub Image32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line12.Visible = True
End Sub

Private Sub Image33_Click()
Clipboard.Clear
Clipboard.SetText Form1.RichTextBox1.SelText, vbCFRTF
Form1.RichTextBox2.Text = ""
Form1.RichTextBox2.TextRTF = Form1.RichTextBox1.SelRTF
Form1.RichTextBox1.SelRTF = ""
End Sub

Private Sub Image33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line13.Visible = True
End Sub

Private Sub Image34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line13.Visible = False
Line12.Visible = False
Line11.Visible = False
End Sub

Private Sub Image36_Click()
SendKeys "^Z"
End Sub

Private Sub Image36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line16.Visible = True
End Sub

Private Sub Image37_Click()
SendKeys "^Y"
End Sub

Private Sub Image37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line17.Visible = True
End Sub

Private Sub Image38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line16.Visible = False
Line17.Visible = False
End Sub

Private Sub Image4_Click()
ShowColorDialogFore
End Sub

Private Sub Image5_Click()
Form1.RichTextBox1.Text = ""
Form1.Caption = "Starboy Notepad"
Form1.Label2.Caption = "Starboy Notepad"
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = True
Line9.Visible = False
Line10.Visible = False
End Sub

Private Sub Image7_Click()
FontOpen
End Sub
Private Sub Image9_Click()
frmOptions.Show
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image29.Visible = True
End Sub


Private Sub Label9_Click()
AddFile
End Sub

Private Sub Text1_Change()
Form1.RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line15.Visible = True
End Sub

Private Sub Text3_Change()
Form1.RichTextBox1.SelFontName = Text3.Text
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Form1.RichTextBox1.SelFontName = Text3.Text
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = Form1.RichTextBox1.SelFontName
Line14.Visible = True

End Sub
