VERSION 4.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5250
   ClientLeft      =   10920
   ClientTop       =   5340
   ClientWidth     =   7395
   Height          =   5655
   Left            =   10860
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Top             =   4995
   Width           =   7515
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      TabIndex        =   48
      Top             =   1560
      Width           =   7215
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FORM5.frx":0000
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   6
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   53
         Top             =   2520
         Width           =   6495
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   6840
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   $"FORM5.frx":009F
         Height          =   375
         Left            =   1080
         TabIndex        =   52
         Top             =   1920
         Width           =   5655
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "For 16-bit Architecture"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   51
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   360
         Picture         =   "FORM5.frx":011A
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.3.6 - build 1360 -a/public-21721 Beta 2"
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Starboy Notepad/16"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   24
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   49
         Top             =   600
         Width           =   4815
      End
      Begin VB.Image Image5 
         Height          =   975
         Left            =   360
         Picture         =   "FORM5.frx":0FE4
         Stretch         =   -1  'True
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   7215
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   5520
         TabIndex        =   46
         Text            =   "Text16"
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   3120
         TabIndex        =   45
         Text            =   "Text15"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox Check17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Background"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox Check16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Windows logo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   40
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox Check15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "16 bit text"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox Check14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "16 bit shape"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tips text"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tips picture"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tips title"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tips line"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tips shape"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Version"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Title shadow"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Title"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Icon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hide"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clear"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6360
         TabIndex        =   47
         Top             =   4560
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Font options"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   6720
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Preview"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   41
         Top             =   3480
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   6720
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enable or disable"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FORM5.frx":18AE
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   6255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C000C0&
         X1              =   360
         X2              =   360
         Y1              =   360
         Y2              =   820
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   7215
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   15
         Text            =   "Verdana"
         Top             =   2520
         Width           =   1500
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   14
         Text            =   "Times New Roman"
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   13
         Text            =   "Tahoma"
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   12
         Text            =   "Segoe UI"
         Top             =   1080
         Width           =   1500
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Text            =   "MS Sans Serif"
         Top             =   3000
         Width           =   1500
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "Lucida Console"
         Top             =   2520
         Width           =   1500
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Text            =   "Impact"
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Text            =   "Haettenschweiler"
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Text            =   "Garamond"
         Top             =   1080
         Width           =   1500
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         ScrollBars      =   1  'Horizontal
         TabIndex        =   6
         Text            =   "Footlight MT Light"
         Top             =   3000
         Width           =   1500
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Text            =   "Courier"
         Top             =   2520
         Width           =   1500
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Text            =   "Bahnschrift"
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "Arial Narrow"
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Text            =   "Arial"
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         X1              =   360
         X2              =   360
         Y1              =   360
         Y2              =   820
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FORM5.frx":190C
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7215
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "About"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   980
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Splash Screen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   980
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Visuals"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   980
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   6120
         Picture         =   "FORM5.frx":1981
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   4200
         Picture         =   "FORM5.frx":224B
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   2280
         Picture         =   "FORM5.frx":3115
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   980
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   360
         Picture         =   "FORM5.frx":3FDF
         Stretch         =   -1  'True
         Top             =   350
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then Form4.Label9.Visible = True
If Check1.Value = 1 Then Form4.Label6.Visible = True
If Check1.Value = 1 Then Form4.Image4.Visible = True
If Check1.Value = 1 Then Form4.Shape2.Visible = True
If Check1.Value = 0 Then Form4.Shape2.Visible = False
If Check1.Value = 0 Then Form4.Label9.Visible = False
If Check1.Value = 0 Then Form4.Label6.Visible = False
If Check1.Value = 0 Then Form4.Image4.Visible = False
If Check1.Value = 0 Then Form4.Shape2.Visible = False
End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then Form4.Line1.Visible = True
If Check10.Value = 0 Then Form4.Line1.Visible = False
End Sub

Private Sub Check11_Click()
If Check11.Value = 1 Then Form4.Label12.Visible = True
If Check11.Value = 0 Then Form4.Label12.Visible = False
End Sub


Private Sub Check13_Click()
If Check13.Value = 1 Then Form4.Label13.Visible = True
If Check13.Value = 0 Then Form4.Label13.Visible = False
End Sub


Private Sub Check14_Click()
If Check14.Value = 1 Then Form4.Shape1.Visible = 1
If Check14.Value = 0 Then Form4.Shape1.Visible = 0
End Sub


Private Sub Check15_Click()
If Check15.Value = 1 Then Form4.Label4.Visible = True
If Check15.Value = 1 Then Form4.Label5.Visible = True
If Check15.Value = 0 Then Form4.Label5.Visible = False
If Check15.Value = 0 Then Form4.Label4.Visible = False

End Sub


Private Sub Check16_Click()
If Check16.Value = 1 Then Form4.Image10.Visible = True
If Check16.Value = 0 Then Form4.Image10.Visible = False
End Sub


Private Sub Check17_Click()
If Check17.Value = 1 Then Form4.Image1.Visible = True
If Check17.Value = 0 Then Form4.Image1.Visible = False
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Form4.Label10.Visible = True
If Check2.Value = 1 Then Form4.Shape3.Visible = True
If Check2.Value = 1 Then Form4.Image3.Visible = True
If Check2.Value = 1 Then Form4.Label7.Visible = True
If Check2.Value = 0 Then Form4.Label10.Visible = False
If Check2.Value = 0 Then Form4.Shape3.Visible = False
If Check2.Value = 0 Then Form4.Image3.Visible = False
If Check2.Value = 0 Then Form4.Label7.Visible = False
End Sub


Private Sub Check3_Click()
If Check3.Value = 1 Then Form4.Label11.Visible = True
If Check3.Value = 1 Then Form4.Label8.Visible = True
If Check3.Value = 1 Then Form4.Shape4.Visible = True
If Check3.Value = 1 Then Form4.Image5.Visible = True
If Check3.Value = 0 Then Form4.Label11.Visible = False
If Check3.Value = 0 Then Form4.Label8.Visible = False
If Check3.Value = 0 Then Form4.Shape4.Visible = False
If Check3.Value = 0 Then Form4.Image5.Visible = False
End Sub


Private Sub Check4_Click()
If Check4.Value = 1 Then Form4.Label18.Visible = True
If Check4.Value = 0 Then Form4.Label18.Visible = False
End Sub


Private Sub Check5_Click()
If Check5.Value = 1 Then Form4.Image2.Visible = True
If Check5.Value = 0 Then Form4.Image2.Visible = False
End Sub


Private Sub Check6_Click()
If Check6.Value = 1 Then Form4.Label2.Visible = True
If Check6.Value = 0 Then Form4.Label2.Visible = False
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then Form4.Label1.Visible = True
If Check7.Value = 0 Then Form4.Label1.Visible = False
End Sub


Private Sub Check8_Click()
If Check8.Value = 1 Then Form4.Label3.Visible = True
If Check8.Value = 0 Then Form4.Label3.Visible = False
End Sub


Private Sub Check9_Click()
If Check9.Value = 1 Then Form4.Shape5.Visible = True
If Check9.Value = 0 Then Form4.Shape5.Visible = False
End Sub


Private Sub Form_Load()
Form5.Left = Form1.Left + 250
Form5.Top = Form1.Top + 250
Frame4.Visible = False
Frame3.Visible = False
Text15.Text = Form4.Label1.Font
Text16.Text = Form4.Label1.Font.Size
If Form4.Image1.Visible = True Then Check17.Value = 1
If Form4.Image1.Visible = False Then Check17.Value = 0
If Form4.Label9.Visible = True Then Check1.Value = 1
If Form4.Label9.Visible = False Then Check1.Value = 0
If Form4.Label10.Visible = True Then Check2.Value = 1
If Form4.Label10.Visible = False Then Check2.Value = 0
If Form4.Label11.Visible = True Then Check3.Value = 1
If Form4.Label11.Visible = False Then Check3.Value = 0
If Form4.Label18.Visible = True Then Check4.Value = 1
If Form4.Label18.Visible = False Then Check4.Value = 0
If Form4.Image2.Visible = True Then Check5.Value = 1
If Form4.Image2.Visible = False Then Check5.Value = 0
If Form4.Label2.Visible = True Then Check6.Value = 1
If Form4.Label2.Visible = False Then Check6.Value = 0
If Form4.Label1.Visible = True Then Check7.Value = 1
If Form4.Label1.Visible = False Then Check7.Value = 0
If Form4.Label3.Visible = True Then Check8.Value = 1
If Form4.Label3.Visible = False Then Check8.Value = 0
If Form4.Shape5.Visible = True Then Check9.Value = 1
If Form4.Shape5.Visible = False Then Check9.Value = 0
If Form4.Line1.Visible = True Then Check10.Value = 1
If Form4.Line1.Visible = False Then Check10.Value = 0
If Form4.Label12.Visible = True Then Check11.Value = 1
If Form4.Label12.Visible = False Then Check11.Value = 0
If Form4.Image6.Visible = True Then Check12.Value = 1
If Form4.Image6.Visible = False Then Check12.Value = 0
If Form4.Label13.Visible = True Then Check13.Value = 1
If Form4.Label13.Visible = False Then Check13.Value = 0
If Form4.Shape1.Visible = True Then Check14.Value = 1
If Form4.Shape1.Visible = False Then Check14.Value = 0
If Form4.Label4.Visible = True Then Check15.Value = 1
If Form4.Label4.Visible = False Then Check15.Value = 0
If Form4.Image10.Visible = True Then Check16.Value = 1
If Form4.Image10.Visible = False Then Check16.Value = 0
End Sub


Private Sub Form_Resize()
Label1.Width = Form5.Width
Label1.Height = Form5.Height
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Font.Italic = False
Label2.Font.Italic = False
Label4.Font.Italic = False
Label6.Font.Italic = False
End Sub


Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)
Label9.BackStyle = 0
Label9.BackStyle = 0
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackStyle = 0
Label12.BackStyle = 0
End Sub


Private Sub Image1_Click()
Label2.Font.Italic = False
Label2.Font.Bold = True
Label5.Font.Bold = False
Label4.Font.Bold = False
Label6.Font.Bold = False
Frame1.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Form5.Height = "5655"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Font.Italic = True
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = True
End Sub


Private Sub Image3_Click()
Label5.Font.Bold = True
Label2.Font.Bold = False
Label4.Font.Bold = False
Label6.Font.Bold = False
Frame1.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Form5.Height = "8040"
End Sub


Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Font.Italic = True
Label2.Font.Italic = False
Label4.Font.Italic = False
Label6.Font.Italic = False
End Sub


Private Sub Image4_Click()
Label6.Font.Bold = True
Label2.Font.Bold = False
Label4.Font.Bold = False
Label5.Font.Bold = False
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Form5.Height = "5190"
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Font.Italic = True
End Sub


Private Sub Label12_Click()
Form4.Label2.Font = Text15.Text
Form4.Label2.Font.Size = Text16.Text
Form4.Label1.Font = Text15.Text
Form4.Label1.Font = Text16.Text
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BackStyle = 1
End Sub


Private Sub Label9_Click()
Form4.Show
End Sub


Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackStyle = 1
End Sub


