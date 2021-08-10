VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8280
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   12165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmSplash.frx":058A
   ScaleHeight     =   8280
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   100
      Left            =   11280
      Top             =   240
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "If you click on this label, you'll enable the debug label in which shows how many seconds have passed."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   1080
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Before this, all the programs I made in Visual Basic were not functional. This one though, luckily is!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   1080
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Did you know this program is all made in Windows XP, Office 2003 and Visual Basic 6.0?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   1080
      TabIndex        =   11
      Top             =   5640
      Width           =   9015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   960
      X2              =   10080
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tip of the day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   2  'Blackness
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   9855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2014-2021 MoonLight Corp. All rights reserved. Enormous thanks to Hanghitorgame for the help in the code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   7920
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   7560
      TabIndex        =   0
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   4080
      TabIndex        =   1
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Image Image5 
      Height          =   585
      Left            =   7920
      Picture         =   "frmSplash.frx":252ADC
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   7680
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   2  'Blackness
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   3195
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   585
      Left            =   4440
      Picture         =   "frmSplash.frx":2531C6
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   2  'Blackness
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   3195
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   840
      Picture         =   "frmSplash.frx":2538B0
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   570
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   2  'Blackness
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3195
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   480
      Picture         =   "frmSplash.frx":253F9A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.53.1545.srv00_FE.090821-0036/32-public"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Starboy Notepad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   6495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Starboy Notepad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1215
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   11010
      Left            =   0
      Picture         =   "frmSplash.frx":254684
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   12570
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sec, min, hrs As Integer

Option Explicit


Private mlngX As Long
Private mlngY As Long


Private Sub Command1_Click()
frmOptions.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Form1.Visible = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = frmSplash.Left
        mlngY = frmSplash.Top
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = RGB(0, 0, 0)
Label8.ForeColor = RGB(0, 0, 0)
Label7.ForeColor = RGB(0, 0, 0)
Label6.BackStyle = 0
Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Label1_Click()

Unload frmSplash
Unload Form1

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = RGB(230, 0, 0)
Label8.ForeColor = RGB(0, 0, 0)
Label7.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Label14_Click()
Unload frmSplash
End Sub

Private Sub Label16_Click()
Label13.Visible = True
End Sub

Private Sub Label2_Click()
Form1.Visible = True
Unload frmSplash
OpenFile
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackStyle = 0
Label9.ForeColor = RGB(0, 0, 0)
Label8.ForeColor = RGB(0, 140, 220)
Label7.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Label3_Click()
Form1.Visible = True
Unload frmSplash
Form1.RichTextBox1.Text = ""
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackStyle = 0
Label9.ForeColor = RGB(0, 0, 0)
Label8.ForeColor = RGB(0, 0, 0)
Label7.ForeColor = RGB(0, 140, 220)
End Sub



Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackStyle = 0
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackStyle = 0
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = RGB(0, 0, 0)
Label8.ForeColor = RGB(0, 0, 0)
Label7.ForeColor = RGB(0, 0, 0)
Label6.BackStyle = 1
End Sub

Private Sub Timer1_Timer(Index As Integer)
sec = sec + 1
Label13.Caption = Format(sec, "00")
If sec = 100 Then Label12.Visible = False
If sec = 100 Then Label15.Visible = True
If sec = 200 Then Label12.Visible = False
If sec = 200 Then Label16.Visible = True
If sec = 200 Then Label15.Visible = False
If sec = 300 Then Label16.Visible = False
If sec = 300 Then Label12.Visible = True
If sec = 300 Then Label15.Visible = False
If sec = 400 Then Label12.Visible = False
If sec = 400 Then Label15.Visible = True
If sec = 400 Then Label16.Visible = False
If sec = 400 Then Label12.Visible = False
End Sub
