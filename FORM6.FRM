VERSION 4.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1050
   ClientLeft      =   23010
   ClientTop       =   5265
   ClientWidth     =   2115
   Height          =   1455
   Left            =   22950
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   Top             =   4920
   Width           =   2235
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Word Count"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   -6120
      TabIndex        =   0
      Top             =   -840
      Width           =   8295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Counts As Integer
If Form1.Text1.Text = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(Form1.Text1.Text)
        If Mid(Form1.Text1.Text, i, 1) = " " Then
            Counts = Counts + 1
        End If
    Next
End If
Text1.Text = Counts
End Sub

Private Sub Form_GotFocus()
Dim Counts As Integer
If Form1.Text1.Text = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(Form1.Text1.Text)
        If Mid(Form1.Text1.Text, i, 1) = " " Then
            Counts = Counts + 1
        End If
    Next
End If
Text1.Text = Counts
End Sub

Private Sub Form_Load()
Form6.Left = Form1.Left + 250
Form6.Top = Form1.Top + 250
Dim Counts As Integer
If Form1.Text1.Text = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(Form1.Text1.Text)
        If Mid(Form1.Text1.Text, i, 1) = " " Then
            Counts = Counts + 1
        End If
    Next
End If
Text1.Text = Counts
End Sub

