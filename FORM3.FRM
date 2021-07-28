VERSION 4.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Starboy Notepad - Colors"
   ClientHeight    =   3855
   ClientLeft      =   9555
   ClientTop       =   4530
   ClientWidth     =   6690
   Height          =   4260
   Left            =   9495
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   4185
   Width           =   6810
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Background"
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   6255
      Begin VB.ComboBox cmbRosso 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbVerde 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbBlue 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblColore 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Font"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   6255
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4680
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   5880
      Top             =   600
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Red"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"FORM3.frx":0000
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   240
      Picture         =   "FORM3.frx":0081
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmbBlue_Change()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

End Sub

Private Sub cmbBlue_GotFocus()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
End Sub


Private Sub cmbBlue_LostFocus()

If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

End Sub

Private Sub cmbRosso_Change()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

End Sub

Private Sub cmbRosso_GotFocus()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
End Sub


Private Sub cmbRosso_LostFocus()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
End Sub


Private Sub cmbVerde_Change()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

End Sub


Private Sub cmbVerde_GotFocus()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
End Sub


Private Sub cmbVerde_LostFocus()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
End Sub
Private Sub Form_Load()
Form3.Left = Form1.Left + 250
Form3.Top = Form1.Top + 250
lblColore.BackColor = Label10.BackColor
lblColore.BorderStyle = 1
Label6.BackColor = Label10.BackColor
Label6.BorderStyle = 1
Dim i As Integer
i = 0
For i = 0 To 255
    Form3.cmbRosso.AddItem (i)
    Form3.cmbVerde.AddItem (i)
   Form3.cmbBlue.AddItem (i)
   Form3.Combo3.AddItem (i)
   Form3.Combo2.AddItem (i)
   Form3.Combo1.AddItem (i)
    Next
    
Frame1.BackColor = Label10.BackColor
Frame2.BackColor = Label10.BackColor


End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

End Sub


Private Sub lblColore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

End Sub


Private Sub Timer1_Timer()
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then lblColore.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then Form1.Text1.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)
If cmbRosso.Text <> "" And cmbVerde.Text <> "" And cmbBlue.Text <> "" Then Form1.Frame1.BackColor = RGB(cmbRosso.Text, cmbVerde.Text, cmbBlue.Text)

If Combo3.Text <> "" And Combo2.Text <> "" And Combo1.Text <> "" Then Label6.BackColor = RGB(Combo3.Text, Combo2.Text, Combo1.Text)
If Combo3.Text <> "" And Combo2.Text <> "" And Combo1.Text <> "" Then Form1.Text1.ForeColor = RGB(Combo3.Text, Combo2.Text, Combo1.Text)
End Sub


