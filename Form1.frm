VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Skrive Lite- Estathè"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   9975
      _Version        =   393216
      TabOrientation  =   2
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "File 1"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "File 2"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "File 3"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image4(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Image5(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Image6(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label2(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text2(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text3(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "CommonDialog1(0)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Combo2(0)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "Form1.frx":0054
         Left            =   1080
         List            =   "Form1.frx":007C
         TabIndex        =   4
         Text            =   "Arial"
         Top             =   215
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Index           =   0
         Left            =   8400
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text3 
         Height          =   4815
         Index           =   0
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   720
         Width           =   10935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Size"
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
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Image Image6 
         Height          =   285
         Index           =   0
         Left            =   9240
         Picture         =   "Form1.frx":0113
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   10320
         Picture         =   "Form1.frx":1C49
         Stretch         =   -1  'True
         Top             =   240
         Width           =   930
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   0
         Left            =   3480
         Picture         =   "Form1.frx":21C6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":A5A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Combo2_GotFocus(Index As Integer)

Text3.
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer, Index As Integer)
Text3.Item.Font = Combo2.Item.Text
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer, Index As Integer)
Text3.Item.Font = Combo2.Item.Text
End Sub






Private Sub Image5_Click(Index As Integer)
OpenFile
End Sub

Private Sub Image6_Click(Index As Integer)
Savefile
End Sub



Private Sub TabStrip1_Click(Index As Integer)

End Sub



Private Sub Text2_Change(Index As Integer)
Text3.Item.Font.Size = Text2.Item.Text
End Sub

Private Sub Text2_GotFocus(Index As Integer)
Text3.Item.Font.Size = Text2.Item.Text
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer, Index As Integer)
Text3.Item.Font.Size = Text2.Item.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer, Index As Integer)
Text3.Item.Font.Size = Text2.Item.Text
End Sub
