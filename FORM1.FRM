VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Starboy Notepad"
   ClientHeight    =   5505
   ClientLeft      =   6720
   ClientTop       =   2385
   ClientWidth     =   9495
   Height          =   6195
   Icon            =   "FORM1.frx":0000
   Left            =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9495
   Top             =   1755
   Width           =   9615
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   64
      Top             =   5040
      Width           =   9495
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   8880
         TabIndex        =   65
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label63 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   195
         Width           =   7455
      End
      Begin VB.Label Label62 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Word count"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7920
         TabIndex        =   66
         Top             =   190
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6360
      TabIndex        =   63
      Text            =   "Combo1"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FORM1.frx":030A
      Left            =   4680
      List            =   "FORM1.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   62
      Text            =   "Combo1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   9495
   End
   Begin WordpadCtl.Document Document1 
      Height          =   285
      Left            =   1680
      OleObjectBlob   =   "FORM1.frx":030E
      TabIndex        =   68
      Top             =   1200
      Width           =   5760
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Font"
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
      Height          =   255
      Left            =   5400
      TabIndex        =   40
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   41
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Underline"
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
      Height          =   255
      Left            =   4800
      TabIndex        =   38
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Italic"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   37
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bold"
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
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FONT"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   4080
      TabIndex        =   20
      Top             =   975
      Width           =   3135
   End
   Begin VB.Label Label51 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   51
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Settings"
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
      Height          =   255
      Left            =   7920
      TabIndex        =   45
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add"
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
      Height          =   255
      Left            =   7440
      TabIndex        =   44
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tools"
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
      Height          =   255
      Left            =   7920
      TabIndex        =   43
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color"
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
      Height          =   255
      Left            =   7440
      TabIndex        =   42
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Font"
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
      Height          =   255
      Left            =   4020
      TabIndex        =   39
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Redo"
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
      Height          =   255
      Left            =   3120
      TabIndex        =   35
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label47 
      Height          =   1215
      Left            =   3720
      TabIndex        =   47
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Paste"
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
      Height          =   255
      Left            =   3120
      TabIndex        =   34
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "U"
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
      Height          =   255
      Left            =   5160
      TabIndex        =   27
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label61 
      Height          =   495
      Left            =   4080
      TabIndex        =   61
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label60 
      Height          =   495
      Left            =   4560
      TabIndex        =   60
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label59 
      Height          =   495
      Left            =   5040
      TabIndex        =   59
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "U"
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
      Height          =   255
      Left            =   5160
      TabIndex        =   26
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label58 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5040
      TabIndex        =   58
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label57 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      TabIndex        =   57
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label56 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      TabIndex        =   56
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   8760
      Picture         =   "FORM1.frx":0D26
      Stretch         =   -1  'True
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label54 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   8520
      TabIndex        =   54
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
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
      Height          =   255
      Left            =   840
      TabIndex        =   30
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copy"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOOLS"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   7560
      TabIndex        =   21
      Top             =   975
      Width           =   975
   End
   Begin VB.Label Label52 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   52
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TEXT ACTIONS"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   1800
      TabIndex        =   19
      Top             =   975
      Width           =   1935
   End
   Begin VB.Label Label50 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   50
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FILE"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   975
      Width           =   1455
   End
   Begin VB.Label Label49 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label48 
      Height          =   1215
      Left            =   7200
      TabIndex        =   48
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label46 
      Height          =   1215
      Left            =   1440
      TabIndex        =   46
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cut"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   32
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Undo"
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
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Open"
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
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
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
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1215
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00F3F3F3&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00F3F3F3&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00F3F3F3&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1440
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1215
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1215
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1215
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   975
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   3135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image14 
      Height          =   255
      Left            =   2880
      Picture         =   "FORM1.frx":1BF0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line9 
      X1              =   6960
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      X1              =   6960
      X2              =   7080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line7 
      X1              =   6840
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      X1              =   6360
      X2              =   6600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line5 
      X1              =   6400
      X2              =   6550
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      X1              =   5880
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   5880
      X2              =   6000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   6120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image13 
      Height          =   255
      Left            =   3360
      Picture         =   "FORM1.frx":217A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   2040
      Picture         =   "FORM1.frx":2704
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   255
      Left            =   7680
      Picture         =   "FORM1.frx":2C8E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   1080
      Picture         =   "FORM1.frx":3218
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   8160
      Picture         =   "FORM1.frx":37A2
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   8180
      Picture         =   "FORM1.frx":3D2C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   7680
      Picture         =   "FORM1.frx":42B6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7560
      TabIndex        =   14
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   4250
      Picture         =   "FORM1.frx":4840
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3375
      Picture         =   "FORM1.frx":4DCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   2900
      Picture         =   "FORM1.frx":5354
      Stretch         =   -1  'True
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   1080
      Picture         =   "FORM1.frx":58DE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   240
      Picture         =   "FORM1.frx":5E68
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label53 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label Label55 
      Caption         =   "Label55"
      Height          =   2175
      Left            =   0
      TabIndex        =   55
      Top             =   120
      Width           =   9375
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu Clear 
         Caption         =   "Clear"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
         Index           =   3
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu SplashScreen 
         Caption         =   "Splash Screen"
         Index           =   5
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
         Index           =   6
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Index           =   7
      Begin VB.Menu cut 
         Caption         =   "Cut"
         Index           =   8
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
         Index           =   9
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
         Index           =   10
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu font 
      Caption         =   "Font"
      Index           =   23
   End
   Begin VB.Menu Alignment 
      Caption         =   "Alignment"
      Index           =   11
      Begin VB.Menu Left 
         Caption         =   "Left"
         Index           =   12
      End
      Begin VB.Menu Center 
         Caption         =   "Center"
         Index           =   13
      End
      Begin VB.Menu Right 
         Caption         =   "Right"
         Index           =   14
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Index           =   18
      Begin VB.Menu Toolswindow 
         Caption         =   "Tools window"
         Index           =   19
         Shortcut        =   ^T
      End
      Begin VB.Menu Addfromfile 
         Caption         =   "Add from file"
         Index           =   21
         Shortcut        =   ^K
      End
      Begin VB.Menu Settings 
         Caption         =   "Settings"
         Index           =   22
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu Colours 
      Caption         =   "Colours"
      Index           =   15
   End
   Begin VB.Menu fuck 
      Caption         =   "Form Popup"
      Visible         =   0   'False
      Begin VB.Menu fontas 
         Caption         =   "Font (F3)"
         Index           =   22
         Shortcut        =   {F3}
      End
      Begin VB.Menu undoa 
         Caption         =   "Undo (CTRL+Z)"
         Index           =   23
      End
      Begin VB.Menu cuta 
         Caption         =   "Cut (CTRL+X)"
         Index           =   24
      End
      Begin VB.Menu copya 
         Caption         =   "Copy (CTRL+C)"
         Index           =   25
      End
      Begin VB.Menu pastea 
         Caption         =   "Paste (CTRL+V)"
         Index           =   26
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Text1.Font = Combo1.Text
End Sub


Private Sub Addfromfile_Click(Index As Integer)
AddFile
End Sub

Private Sub Clear_Click(Index As Integer)
If Text1.Text = "" Then Text1.Text = ""
If Not Text1.Text = "" Then Form2.Show

Form1.Caption = "Starboy Notepad"
End Sub

Private Sub Close_Click(Index As Integer)
If Not Text1.Text = "" Then Form2.Show
If Text1.Text = "" Then Unload Form1

End Sub


Private Sub Colours_Click(Index As Integer)
Form3.Show
End Sub

Private Sub Combo1_Change()
Text1.Font = Combo1.Text
End Sub

Private Sub Combo1_GotFocus()
Text1.Font = Combo1.Text
End Sub


Private Sub Combo2_Change()
Text1.Font.Size = Combo2.Text
End Sub

Private Sub Combo2_Click()
Text1.Font.Size = Combo2.Text
End Sub


Private Sub Copy_Click(Index As Integer)
SendKeys "^C"
End Sub

Private Sub Cut_Click(Index As Integer)
SendKeys "^X"
End Sub

Private Sub Font_Click(Index As Integer)
FontOpen
End Sub

Private Sub fontas_Click(Index As Integer)
FontOpen
End Sub


Private Sub Form_GotFocus()
Combo1.Text = Text1.Font
Combo2.Text = Text1.Font.Size
End Sub

Private Sub Form_Load()
Form4.Show
Form1.Hide
Label59.Visible = False
Label60.Visible = False
Label61.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label23.Visible = False
Label25.Visible = False
Label27.Visible = False
Label27.Font.Underline = True
Label26.Font.Underline = True
Combo1.Text = Text1.Font
Combo2.Text = Text1.Font.Size
Form1.Combo1.AddItem (Form5.Text1.Text)
Form1.Combo1.AddItem (Form5.Text2.Text)
Form1.Combo1.AddItem (Form5.Text3.Text)
Form1.Combo1.AddItem (Form5.Text4.Text)
Form1.Combo1.AddItem (Form5.Text5.Text)
Form1.Combo1.AddItem (Form5.Text6.Text)
Form1.Combo1.AddItem (Form5.Text7.Text)
Form1.Combo1.AddItem (Form5.Text8.Text)
Form1.Combo1.AddItem (Form5.Text9.Text)
Form1.Combo1.AddItem (Form5.Text10.Text)
Form1.Combo1.AddItem (Form5.Text11.Text)
Form1.Combo1.AddItem (Form5.Text12.Text)
Form1.Combo1.AddItem (Form5.Text13.Text)
Form1.Combo1.AddItem (Form5.Text14.Text)
Form1.Combo2.AddItem "4"
Form1.Combo2.AddItem "6"
Form1.Combo2.AddItem "8"
Form1.Combo2.AddItem "10"
Form1.Combo2.AddItem "12"
Form1.Combo2.AddItem "14"
Form1.Combo2.AddItem "16"
Form1.Combo2.AddItem "18"
Form1.Combo2.AddItem "20"
Form1.Combo2.AddItem "24"
Form1.Combo2.AddItem "28"
Form1.Combo2.AddItem "36"
Form1.Combo2.AddItem "48"
Form1.Combo2.AddItem "52"
Form1.Combo2.AddItem "64"
Form1.Combo2.AddItem "72"

End Sub


Private Sub SSCommand6_Click()

End Sub


Private Sub Form_Resize()
Text1.Width = Form1.Width
Text1.Height = Form1.Height
Label53.Width = Form1.Width
Label55.Width = Form1.Width
Label55.Height = Form1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form4
Unload Form3
Unload Form5
Unload Form6
End Sub

Private Sub Image1_Click(Index As Integer)
If Text1.Text = "" Then Text1.Text = ""
If Not Text1.Text = "" Then Form2.Show

Form1.Caption = "Starboy Notepad"
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackStyle = 0
Label28.Visible = True
End Sub


Private Sub Image11_Click()
Form3.Show
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.BackStyle = 0
Label42.Visible = True
End Sub


Private Sub Image12_Click()
SendKeys "^Z"
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackStyle = 0
Label31.Visible = True
End Sub


Private Sub Image13_Click()
SendKeys "^Y"
End Sub

Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackStyle = 0
Label34.Visible = False
Label35.Visible = True
End Sub


Private Sub Image14_Click()
SendKeys "^X"
End Sub

Private Sub Image14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackStyle = 0
Label32.Visible = True
End Sub


Private Sub Image2_Click()
OpenFile
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackStyle = 0
Label29.Visible = True
End Sub


Private Sub Image3_Click()
SendKeys "^V"
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackStyle = 0
Label34.Visible = True
End Sub


Private Sub Image4_Click()
SendKeys "^C"
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackStyle = 0
Label32.Visible = False
Label33.Visible = True
End Sub


Private Sub Image5_Click()
FontOpen
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label39.Visible = True
End Sub


Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.BackStyle = 0
Label43.Visible = True
End Sub


Private Sub Image7_Click()
Form5.Show
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.BackStyle = 0
Label43.Visible = False
Label45.Visible = True
End Sub


Private Sub Image8_Click()
AddFile
End Sub


Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.BackStyle = 0
Label42.Visible = False
Label44.Visible = True
End Sub


Private Sub Image9_Click()
salva_bxy
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 0
Label29.Visible = False
Label30.Visible = True
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 1
Label1.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.BackStyle = 0
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label22_Click()
Label61.Visible = True
Text1.Font.Bold = True
Label23.Visible = True
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.Visible = True
Label39.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub


Private Sub Label23_Click()
Label61.Visible = False
Text1.Font.Bold = False
Label23.Visible = False
End Sub


Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.Visible = True
Label39.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub


Private Sub Label24_Click()
Label60.Visible = True
Label24.Visible = True
Label25.Visible = True
Text1.Font.Italic = True
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.Visible = False
Label37.Visible = True
Label38.Visible = False
End Sub


Private Sub Label25_Click()
Label60.Visible = False
Label24.Visible = True
Label25.Visible = False
Text1.Font.Italic = False
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.Visible = False
Label37.Visible = True
Label38.Visible = False
End Sub


Private Sub Label26_Click()
Label59.Visible = True
Label27.Visible = True
Text1.Font.Underline = True
End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label37.Visible = False
Label36.Visible = False
Label38.Visible = True
End Sub


Private Sub Label27_Click()
Label59.Visible = False
Label27.Visible = False

Text1.Font.Underline = False
End Sub

Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label37.Visible = False
Label36.Visible = False
Label38.Visible = True
End Sub


Private Sub Label28_Click()
Label3.BackStyle = 1
Label1.BackStyle = 1
End Sub

Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label2.BackStyle = 1
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 1
End Sub


Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackStyle = 1
End Sub


Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False

Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackStyle = 1
End Sub


Private Sub Label34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False

Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackStyle = 1
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.BackStyle = 1
End Sub


Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.BackStyle = 1
End Sub


Private Sub Label44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.BackStyle = 1
End Sub


Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.BackStyle = 1
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label6_Click()
Label6.BackStyle = 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackStyle = 1
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Open_Click(Index As Integer)
OpenFile
End Sub

Private Sub Paste_Click(Index As Integer)
SendKeys "^V"
End Sub

Private Sub Save_Click(Index As Integer)
salva_bxy
End Sub

Private Sub SSCommand11_Click()
If Text1.Font.Bold = True Then Text1.Font.Bold = False
If Text1.Font.Bold = False Then Text1.Font.Bold = True
End Sub

Private Sub SSCommand12_Click()
If Text1.Font.Italic = True Then Text1.Font.Italic = False
If Text1.Font.Italic = False Then Text1.Font.Italic = True
End Sub

Private Sub SSCommand13_Click()
If Text1.Font.Underline = True Then Text1.Font.Underline = False
If Text1.Font.Underline = False Then Text1.Font.Underline = True
If Text1.Font.Underline = False Then SSCommand13.Font3D = 0
If Text1.Font.Underline = True Then SSCommand13.Font3D = 1
End Sub


Private Sub Settings_Click(Index As Integer)
Form5.Show
End Sub

Private Sub Text1_Change()
On Error Resume Next
Dim i As Integer
Dim Mid As Integer
If Text1.Font.Bold = True Then Label23.Visible = True

If Text1.Font.Italic = True Then Label25.Visible = True

If Text1.Font.Underline = True Then Label27.Visible = True



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
Text2.Text = Counts
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Font = Combo1.Text
Combo2.Text = Text1.Font.Size
If Text1.Font.Bold = True Then Label23.Visible = True

If Text1.Font.Italic = True Then Label25.Visible = True

If Text1.Font.Underline = True Then Label27.Visible = True

Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
End Sub


Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu fuck
End If
End Sub

Private Sub Text2_Change()
Text1.Font = Text2.Text
End Sub


Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label40.Visible = True
End Sub


Private Sub Text3_Change()
Text1.Font.Size = Text3.Text
End Sub


Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label41.Visible = True
End Sub




Private Sub WordWrap_Click(Index As Integer)

End Sub


Private Sub wordcount_Click(Index As Integer)
Form6.Show
End Sub


