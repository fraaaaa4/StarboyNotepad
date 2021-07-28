VERSION 4.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8250
   ClientLeft      =   5175
   ClientTop       =   7845
   ClientWidth     =   12915
   Height          =   8655
   Left            =   5115
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FORM12.frx":0000
   ScaleHeight     =   8.618
   ScaleMode       =   0  'User
   ScaleWidth      =   36.697
   ShowInTaskbar   =   0   'False
   Top             =   7500
   Width           =   13035
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hide"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "0"
      Height          =   375
      Left            =   10920
      TabIndex        =   16
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   $"FORM12.frx":6C0042
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   15
      Top             =   5520
      Width           =   8055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   $"FORM12.frx":6C00E8
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   14
      Top             =   5520
      Width           =   8055
   End
   Begin VB.Image Image6 
      Height          =   573
      Left            =   1320
      Picture         =   "FORM12.frx":6C019C
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   $"FORM12.frx":6C04A6
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   13
      Top             =   5520
      Width           =   8055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"FORM12.frx":6C0513
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   12
      Top             =   5520
      Width           =   8055
   End
   Begin VB.Line Line1 
      X1              =   20.183
      X2              =   157.798
      Y1              =   123.519
      Y2              =   123.519
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Tip of the day:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   3  'Not Merge Pen
      Height          =   2895
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   10095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   10
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4560
      TabIndex        =   9
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   960
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   8400
      Picture         =   "FORM12.frx":6C05D1
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   4800
      Picture         =   "FORM12.frx":6C149B
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   1200
      Picture         =   "FORM12.frx":6C1A25
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      Height          =   1215
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      Height          =   1215
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      Height          =   1215
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "bit"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   11465
      TabIndex        =   4
      Top             =   1278
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   3
      Top             =   1158
      Width           =   375
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   10320
      Picture         =   "FORM12.frx":6C22EF
      Stretch         =   -1  'True
      Top             =   1128
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      Height          =   855
      Left            =   10080
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.3.4 - build 1345 -a/public-21720 - 16-bit Edition"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Starboy Notepad"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   48
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   1
      Top             =   780
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Starboy Notepad"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   48
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin VB.Image Image2 
      Height          =   1380
      Left            =   600
      Picture         =   "FORM12.frx":6C31B9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   8295
      Left            =   0
      Picture         =   "FORM12.frx":6C3A83
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   -120
      TabIndex        =   18
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Secondi As Integer

Private Sub Form_Load()
Label17.Visible = False
Label16.Visible = False
Label15.Visible = False
Label14.Visible = False


Form4.Left = (Screen.Width \ 2) - (Form4.Width \ 2)
Form4.Top = (Screen.Height \ 2) - (Form4.Height \ 2)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = RGB(255, 255, 255)
Label7.ForeColor = RGB(255, 255, 255)
Label8.ForeColor = RGB(255, 255, 255)
Label18.BackStyle = 0
End Sub

Private Sub Label10_Click()
OpenFile
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = RGB(40, 196, 40)
End Sub


Private Sub Label11_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4

End Sub


Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = RGB(236, 55, 55)
End Sub


Private Sub Label14_Click()
Label17.Visible = True
End Sub

Private Sub Label17_Click()
Label17.Visible = False
End Sub

Private Sub Label18_Click()
Form4.Hide
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.BackStyle = 1
End Sub


Private Sub Label9_Click()
Unload Form4
Form1.Show
End Sub



Private Sub Timer1_Timer()
Label17.Caption = Label17.Caption + 1
If Label17.Caption = "0" Then Label14.Visible = False
If Label17.Caption = "0" Then Label15.Visible = False
If Label17.Caption = "0" Then Label16.Visible = False
If Label17.Caption = "0" Then Label13.Visible = True
If Label17.Caption = "150" Then Label14.Visible = True
If Label17.Caption = "150" Then Label115.Visible = False
If Label17.Caption = "150" Then Label116.Visible = False
If Label17.Caption = "150" Then Label113.Visible = False
If Label17.Caption = "300" Then Label15.Visible = True
If Label17.Caption = "300" Then Label14.Visible = False
If Label17.Caption = "300" Then Label16.Visible = False
If Label17.Caption = "300" Then Label13.Visible = False
If Label17.Caption = "450" Then Label16.Visible = True
If Label17.Caption = "450" Then Label15.Visible = False
If Label17.Caption = "450" Then Label14.Visible = False
If Label17.Caption = "450" Then Label13.Visible = False
If Label17.Caption = "600" Then Label13.Visible = True
If Label17.Caption = "600" Then Label14.Visible = False
If Label17.Caption = "600" Then Label15.Visible = False
If Label17.Caption = "600" Then Label16.Visible = False
If Label17.Caption = "750" Then Label14.Visible = True
If Label17.Caption = "750" Then Label15.Visible = False
If Label17.Caption = "750" Then Label16.Visible = False
If Label17.Caption = "750" Then Label13.Visible = False
If Label17.Caption = "900" Then Label15.Visible = True
If Label17.Caption = "900" Then Label14.Visible = False
If Label17.Caption = "900" Then Label16.Visible = False
If Label17.Caption = "900" Then Label13.Visible = False
If Label17.Caption = "1050" Then Label16.Visible = True
If Label17.Caption = "1050" Then Label15.Visible = False
If Label17.Caption = "1050" Then Label14.Visible = False
If Label17.Caption = "1050" Then Label13.Visible = False
If Label17.Caption = "1200" Then Label13.Visible = True
If Label17.Caption = "1200" Then Label14.Visible = False
If Label17.Caption = "1200" Then Label15.Visible = False
If Label17.Caption = "1200" Then Label16.Visible = False
If Label17.Caption = "1350" Then Label14.Visible = True
If Label17.Caption = "1350" Then Label15.Visible = False
If Label17.Caption = "1350" Then Label16.Visible = False
If Label17.Caption = "1350" Then Label13.Visible = False
If Label17.Caption = "1500" Then Label15.Visible = True
If Label17.Caption = "1500" Then Label14.Visible = False
If Label17.Caption = "1500" Then Label16.Visible = False
If Label17.Caption = "1500" Then Label13.Visible = False
If Label17.Caption = "1650" Then Label16.Visible = True
If Label17.Caption = "1650" Then Label15.Visible = False
If Label17.Caption = "1650" Then Label14.Visible = False
If Label17.Caption = "1650" Then Label13.Visible = False
If Label17.Caption = "1800" Then Label13.Visible = True
If Label17.Caption = "1800" Then Label14.Visible = False
If Label17.Caption = "1800" Then Label15.Visible = False
If Label17.Caption = "1800" Then Label16.Visible = False
If Label17.Caption = "1950" Then Label14.Visible = True
If Label17.Caption = "1950" Then Label15.Visible = False
If Label17.Caption = "1950" Then Label16.Visible = False
If Label17.Caption = "1950" Then Label13.Visible = False
If Label17.Caption = "2100" Then Label15.Visible = True
If Label17.Caption = "2100" Then Label14.Visible = False
If Label17.Caption = "2100" Then Label16.Visible = False
If Label17.Caption = "2100" Then Label13.Visible = False
If Label17.Caption = "2250" Then Label16.Visible = True
If Label17.Caption = "2250" Then Label15.Visible = False
If Label17.Caption = "2250" Then Label14.Visible = False
If Label17.Caption = "2250" Then Label13.Visible = False
If Label17.Caption = "2400" Then Label13.Visible = True
If Label17.Caption = "2400" Then Label14.Visible = False
If Label17.Caption = "2400" Then Label15.Visible = False
If Label17.Caption = "2400" Then Label16.Visible = False
If Label17.Caption = "2550" Then Label14.Visible = True
If Label17.Caption = "2550" Then Label15.Visible = False
If Label17.Caption = "2550" Then Label16.Visible = False
If Label17.Caption = "2550" Then Label13.Visible = False
If Label17.Caption = "2700" Then Label15.Visible = True
If Label17.Caption = "2700" Then Label14.Visible = False
If Label17.Caption = "2700" Then Label16.Visible = False
If Label17.Caption = "2700" Then Label13.Visible = False
End Sub


Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = RGB(5, 188, 226)
End Sub


