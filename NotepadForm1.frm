VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Starboy Notepad"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12270
   Icon            =   "NotepadForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   495
      Left            =   12720
      TabIndex        =   48
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"NotepadForm1.frx":058A
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "The controls in this window have been moved to the Tools window. Once you resize, they'll get restored."
      Top             =   0
      Width           =   12255
   End
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
      Left            =   22200
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17640
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5415
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9551
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"NotepadForm1.frx":0615
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
      Height          =   195
      Left            =   3840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   450
      Width           =   1335
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
      Height          =   195
      Left            =   3840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   795
      Width           =   495
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   2760
      TabIndex        =   36
      ToolTipText     =   "CTRL+Z"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image80 
      Height          =   255
      Left            =   2400
      Picture         =   "NotepadForm1.frx":06A0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image79 
      Height          =   255
      Left            =   2400
      Picture         =   "NotepadForm1.frx":DACE
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   2760
      TabIndex        =   35
      ToolTipText     =   "CTRL+C"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image78 
      Height          =   255
      Left            =   2400
      Picture         =   "NotepadForm1.frx":1B800
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   2760
      TabIndex        =   34
      ToolTipText     =   "CTRL+V"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape Shape14 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   1695
      Left            =   2280
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   1695
      Left            =   2400
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image82 
      Height          =   255
      Left            =   4200
      Picture         =   "NotepadForm1.frx":28C2E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image81 
      Height          =   255
      Left            =   4200
      Picture         =   "NotepadForm1.frx":334B4
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   4560
      TabIndex        =   42
      ToolTipText     =   "CTRL+Z"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   4560
      TabIndex        =   41
      ToolTipText     =   "CTRL+C"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape16 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   1215
      Left            =   4080
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   1215
      Left            =   4200
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image76 
      Height          =   255
      Left            =   3480
      Picture         =   "NotepadForm1.frx":524A6
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   3840
      TabIndex        =   40
      ToolTipText     =   "CTRL+V"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Image Image75 
      Height          =   255
      Left            =   3480
      Picture         =   "NotepadForm1.frx":62910
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image74 
      Height          =   255
      Left            =   3480
      Picture         =   "NotepadForm1.frx":771AA
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Image73 
      Height          =   255
      Left            =   3480
      Picture         =   "NotepadForm1.frx":87614
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Bullets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   3840
      TabIndex        =   39
      ToolTipText     =   "CTRL+Z"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      ToolTipText     =   "CTRL+C"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   3840
      TabIndex        =   37
      ToolTipText     =   "CTRL+V"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2175
      Left            =   3360
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2175
      Left            =   3480
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Paste"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   46
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   45
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Cut"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   44
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   13740
      TabIndex        =   43
      Top             =   795
      Width           =   1935
   End
   Begin VB.Label Label43 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   13800
      TabIndex        =   47
      Top             =   840
      Width           =   1935
   End
   Begin VB.Image Image72 
      Height          =   255
      Left            =   1080
      Picture         =   "NotepadForm1.frx":9D256
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Redo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      ToolTipText     =   "CTRL+Y"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Undo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   1440
      TabIndex        =   32
      ToolTipText     =   "CTRL+Z"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Paste"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   1440
      TabIndex        =   31
      ToolTipText     =   "CTRL+V"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      ToolTipText     =   "CTRL+C"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Cut"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      ToolTipText     =   "CTRL+Z"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image71 
      Height          =   255
      Left            =   1080
      Picture         =   "NotepadForm1.frx":CC708
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   1080
      X2              =   3000
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image70 
      Height          =   255
      Left            =   1080
      Picture         =   "NotepadForm1.frx":FB1C6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image69 
      Height          =   255
      Left            =   1080
      Picture         =   "NotepadForm1.frx":10C5A8
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Image68 
      Height          =   255
      Left            =   1080
      Picture         =   "NotepadForm1.frx":11C436
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2895
      Left            =   960
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Splash Screen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image67 
      Height          =   255
      Left            =   600
      Picture         =   "NotepadForm1.frx":12BE3C
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   600
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image66 
      Height          =   255
      Left            =   600
      Picture         =   "NotepadForm1.frx":139D9E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image65 
      Height          =   255
      Left            =   600
      Picture         =   "NotepadForm1.frx":144538
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Image64 
      Height          =   255
      Left            =   600
      Picture         =   "NotepadForm1.frx":14E39A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2415
      Left            =   480
      Top             =   360
      Width           =   2055
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2415
      Left            =   600
      Top             =   480
      Width           =   2010
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Colours"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   75
      Width           =   735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
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
      Left            =   3360
      TabIndex        =   23
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment"
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
      Left            =   2280
      TabIndex        =   22
      Top             =   75
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
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
      Left            =   960
      TabIndex        =   21
      Top             =   75
      Width           =   375
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
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
      Left            =   480
      TabIndex        =   20
      Top             =   75
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "New style"
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
      Left            =   4920
      TabIndex        =   19
      Top             =   75
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Colours"
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
      Left            =   4080
      TabIndex        =   18
      Top             =   75
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
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
      Left            =   3360
      TabIndex        =   17
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Alignment"
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
      Left            =   2280
      TabIndex        =   16
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      Left            =   1560
      TabIndex        =   15
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
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
      Left            =   960
      TabIndex        =   14
      Top             =   75
      Width           =   375
   End
   Begin VB.Image Image63 
      Height          =   255
      Left            =   60
      Picture         =   "NotepadForm1.frx":1872F8
      Stretch         =   -1  'True
      Top             =   45
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
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
      TabIndex        =   13
      Top             =   75
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Starboy Notepad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   80
      Width           =   2415
   End
   Begin VB.Image Image46 
      Height          =   210
      Left            =   11190
      Picture         =   "NotepadForm1.frx":187882
      Stretch         =   -1  'True
      Top             =   70
      Width           =   210
   End
   Begin VB.Image Image45 
      Height          =   220
      Left            =   11550
      Picture         =   "NotepadForm1.frx":18F6A4
      Stretch         =   -1  'True
      Top             =   70
      Width           =   220
   End
   Begin VB.Image Image44 
      Height          =   220
      Left            =   11915
      Picture         =   "NotepadForm1.frx":198746
      Stretch         =   -1  'True
      Top             =   70
      Width           =   220
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   50
      Width           =   300
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   300
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   50
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   50
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   300
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image43 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":1A0C6C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image42 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":1A50A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bullets"
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
      Left            =   7680
      TabIndex        =   10
      Top             =   465
      Width           =   855
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00C0C000&
      X1              =   2640
      X2              =   2850
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C0C000&
      X1              =   1680
      X2              =   2040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image Image36 
      Height          =   435
      Left            =   1620
      Picture         =   "NotepadForm1.frx":1A94E0
      Stretch         =   -1  'True
      ToolTipText     =   "Undo (CTRL+Z)"
      Top             =   510
      Width           =   435
   End
   Begin VB.Image Image35 
      Height          =   330
      Left            =   11640
      Picture         =   "NotepadForm1.frx":1D7F9E
      Stretch         =   -1  'True
      Top             =   915
      Width           =   495
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0C000&
      X1              =   3840
      X2              =   4320
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C000&
      X1              =   3840
      X2              =   5160
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0C000&
      X1              =   2265
      X2              =   2475
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C000&
      X1              =   2280
      X2              =   2490
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C000&
      X1              =   2640
      X2              =   2850
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C000&
      X1              =   840
      X2              =   1050
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C000&
      X1              =   840
      X2              =   1050
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C000&
      X1              =   240
      X2              =   720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image Image25 
      Height          =   195
      Left            =   840
      Picture         =   "NotepadForm1.frx":1FD7FC
      Stretch         =   -1  'True
      ToolTipText     =   "Saves the current file"
      Top             =   795
      Width           =   195
   End
   Begin VB.Image Image13 
      Height          =   195
      Left            =   840
      Picture         =   "NotepadForm1.frx":207F96
      Stretch         =   -1  'True
      ToolTipText     =   "Opens a new file"
      Top             =   450
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   435
      Left            =   255
      Picture         =   "NotepadForm1.frx":211DF8
      Stretch         =   -1  'True
      ToolTipText     =   "Cleans the existing file. Note: all unsaved changes will be discarded"
      Top             =   510
      Width           =   435
   End
   Begin VB.Image Image30 
      Height          =   1080
      Left            =   120
      Picture         =   "NotepadForm1.frx":24AD56
      Stretch         =   -1  'True
      Top             =   315
      Width           =   1125
   End
   Begin VB.Image Image9 
      Height          =   330
      Left            =   11160
      Picture         =   "NotepadForm1.frx":2D3D94
      Stretch         =   -1  'True
      Top             =   915
      Width           =   345
   End
   Begin VB.Image Image27 
      Height          =   330
      Left            =   11040
      Picture         =   "NotepadForm1.frx":2E41FE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image29 
      Height          =   330
      Left            =   11160
      Picture         =   "NotepadForm1.frx":32DB48
      Stretch         =   -1  'True
      Top             =   915
      Width           =   345
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11040
      TabIndex        =   8
      Top             =   840
      Width           =   855
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
      Left            =   21000
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image28 
      Height          =   330
      Left            =   22080
      Picture         =   "NotepadForm1.frx":33D14E
      Stretch         =   -1  'True
      Top             =   435
      Width           =   495
   End
   Begin VB.Image Image21 
      Height          =   210
      Left            =   7440
      Picture         =   "NotepadForm1.frx":3A2390
      Stretch         =   -1  'True
      ToolTipText     =   "Adds text bullets to your current row"
      Top             =   450
      Width           =   255
   End
   Begin VB.Image Image26 
      Height          =   210
      Left            =   7440
      Picture         =   "NotepadForm1.frx":3B6C2A
      Stretch         =   -1  'True
      ToolTipText     =   "Removes text bullets from your current row"
      Top             =   450
      Width           =   255
   End
   Begin VB.Image Image12 
      Height          =   345
      Left            =   6120
      Picture         =   "NotepadForm1.frx":3CB5CC
      Stretch         =   -1  'True
      ToolTipText     =   "Center alignment"
      Top             =   555
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   345
      Left            =   6600
      Picture         =   "NotepadForm1.frx":3D886E
      Stretch         =   -1  'True
      ToolTipText     =   "Right alignment"
      Top             =   555
      Width           =   345
   End
   Begin VB.Image Image10 
      Height          =   345
      Left            =   5640
      Picture         =   "NotepadForm1.frx":3E5C9C
      Stretch         =   -1  'True
      ToolTipText     =   "Left alignment"
      Top             =   555
      Width           =   345
   End
   Begin VB.Image Image24 
      Height          =   345
      Left            =   5640
      Picture         =   "NotepadForm1.frx":3F30CA
      Stretch         =   -1  'True
      Top             =   555
      Width           =   345
   End
   Begin VB.Image Image23 
      Height          =   345
      Left            =   6600
      Picture         =   "NotepadForm1.frx":40026C
      Stretch         =   -1  'True
      Top             =   555
      Width           =   345
   End
   Begin VB.Image Image22 
      Height          =   345
      Left            =   6120
      Picture         =   "NotepadForm1.frx":40D0EE
      Stretch         =   -1  'True
      Top             =   555
      Width           =   345
   End
   Begin VB.Image Image19 
      Height          =   270
      Left            =   4920
      Picture         =   "NotepadForm1.frx":41A63C
      Stretch         =   -1  'True
      ToolTipText     =   "Underline"
      Top             =   795
      Width           =   270
   End
   Begin VB.Image Image17 
      Height          =   270
      Left            =   4680
      Picture         =   "NotepadForm1.frx":4276BE
      Stretch         =   -1  'True
      ToolTipText     =   "Italic"
      Top             =   795
      Width           =   270
   End
   Begin VB.Image Image15 
      Height          =   270
      Left            =   4440
      Picture         =   "NotepadForm1.frx":4352F0
      Stretch         =   -1  'True
      ToolTipText     =   "Bold"
      Top             =   795
      Width           =   270
   End
   Begin VB.Image Image20 
      Height          =   270
      Left            =   4920
      Picture         =   "NotepadForm1.frx":442A5A
      Stretch         =   -1  'True
      ToolTipText     =   "Underline"
      Top             =   795
      Width           =   270
   End
   Begin VB.Image Image18 
      Height          =   270
      Left            =   4680
      Picture         =   "NotepadForm1.frx":44FE14
      Stretch         =   -1  'True
      ToolTipText     =   "Italic"
      Top             =   795
      Width           =   270
   End
   Begin VB.Image Image16 
      Height          =   270
      Left            =   4440
      Picture         =   "NotepadForm1.frx":45D1CE
      Stretch         =   -1  'True
      ToolTipText     =   "Bold"
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   7680
      TabIndex        =   3
      Top             =   795
      Width           =   1095
   End
   Begin VB.Image Image14 
      Height          =   225
      Left            =   7455
      Picture         =   "NotepadForm1.frx":46A9B8
      Stretch         =   -1  'True
      ToolTipText     =   "Adds text to this file from another file"
      Top             =   795
      Width           =   225
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
      Left            =   9840
      TabIndex        =   2
      Top             =   810
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   10320
      Picture         =   "NotepadForm1.frx":4805FA
      Stretch         =   -1  'True
      ToolTipText     =   "Changes the colour of the selected font"
      Top             =   765
      Width           =   390
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   10320
      Picture         =   "NotepadForm1.frx":48DABC
      Stretch         =   -1  'True
      ToolTipText     =   "Changes the colour of the background"
      Top             =   420
      Width           =   390
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
      Left            =   9240
      TabIndex        =   1
      Top             =   465
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   3720
      Picture         =   "NotepadForm1.frx":49AF7E
      Stretch         =   -1  'True
      Top             =   765
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3720
      Picture         =   "NotepadForm1.frx":5001C0
      Stretch         =   -1  'True
      Top             =   405
      Width           =   1485
   End
   Begin VB.Image Image39 
      Height          =   1080
      Left            =   5520
      Picture         =   "NotepadForm1.frx":565402
      Stretch         =   -1  'True
      Top             =   315
      Width           =   1575
   End
   Begin VB.Image Image8 
      Height          =   1080
      Left            =   7320
      Picture         =   "NotepadForm1.frx":60A384
      Stretch         =   -1  'True
      Top             =   315
      Width           =   1575
   End
   Begin VB.Image Image40 
      Height          =   1080
      Left            =   9120
      Picture         =   "NotepadForm1.frx":6B57DA
      Stretch         =   -1  'True
      Top             =   315
      Width           =   1695
   End
   Begin VB.Image Image47 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":771E74
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image57 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":777A86
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image56 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":77D698
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image55 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":780FAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image54 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":7871A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image53 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":78D2C5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image50 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":792B3A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image48 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":798BC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image49 
      Height          =   360
      Left            =   -120
      Picture         =   "NotepadForm1.frx":79E157
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image51 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":7A36E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image52 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":7A9391
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image60 
      Height          =   360
      Left            =   -240
      Picture         =   "NotepadForm1.frx":7AF03A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image59 
      Height          =   360
      Left            =   -360
      Picture         =   "NotepadForm1.frx":7B515D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image58 
      Height          =   360
      Left            =   -360
      Picture         =   "NotepadForm1.frx":7BA9D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image62 
      Height          =   360
      Left            =   -120
      Picture         =   "NotepadForm1.frx":7C0A5E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image61 
      Height          =   360
      Left            =   -120
      Picture         =   "NotepadForm1.frx":7C4374
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   2895
      Left            =   1080
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image31 
      Height          =   195
      Left            =   2640
      Picture         =   "NotepadForm1.frx":7CA568
      Stretch         =   -1  'True
      ToolTipText     =   "Paste (CTRL+V)"
      Top             =   450
      Width           =   180
   End
   Begin VB.Image Image33 
      Height          =   195
      Left            =   2280
      Picture         =   "NotepadForm1.frx":7DB94A
      Stretch         =   -1  'True
      ToolTipText     =   "Cut (CTRL+Z)"
      Top             =   450
      Width           =   195
   End
   Begin VB.Image Image32 
      Height          =   195
      Left            =   2280
      Picture         =   "NotepadForm1.frx":7EB350
      Stretch         =   -1  'True
      ToolTipText     =   "Copy (CTRL+C)"
      Top             =   795
      Width           =   195
   End
   Begin VB.Image Image37 
      Height          =   195
      Left            =   2640
      Picture         =   "NotepadForm1.frx":7FB1DE
      Stretch         =   -1  'True
      ToolTipText     =   "Redo (CTRL+Y)"
      Top             =   795
      Width           =   180
   End
   Begin VB.Image Image34 
      Height          =   1080
      Left            =   1440
      Picture         =   "NotepadForm1.frx":82A690
      Stretch         =   -1  'True
      Top             =   315
      Width           =   1635
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   3360
      Picture         =   "NotepadForm1.frx":8E34D2
      Stretch         =   -1  'True
      Top             =   795
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   225
      Left            =   3360
      Picture         =   "NotepadForm1.frx":902360
      Stretch         =   -1  'True
      ToolTipText     =   "Change the font of your text"
      Top             =   450
      Width           =   255
   End
   Begin VB.Image Image38 
      Height          =   1080
      Left            =   3300
      Picture         =   "NotepadForm1.frx":921352
      Stretch         =   -1  'True
      ToolTipText     =   "Change the font of your text"
      Top             =   315
      Width           =   2055
   End
   Begin VB.Image Image41 
      Height          =   1200
      Left            =   0
      Picture         =   "NotepadForm1.frx":A004A4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   12315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mlngX As Long
Private mlngY As Long

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF


' Show the common dialog for choosing a color.
' Return the chosen color, or -1 if the dialog is canceled
'
' hParent is the handle of the parent form
' bFullOpen specifies whether the dialog will be open with the Full style
' (allows to choose many more colors)
' InitColor is the color initially selected when the dialog is open

' Example:
'    Dim oleNewColor As OLE_COLOR
'    oleNewColor = ShowColorsDialog(Me.hwnd, True, vbRed)
'    If oleNewColor <> -1 Then Me.BackColor = oleNewColor

Function ShowColorDialog(Optional ByVal hParent As Long, _
    Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) _
    As Long
    Dim CC As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long

    ' translate the initial OLE color to a long value
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    'fill the ChooseColorStruct struct
    With CC
        .lStructSize = Len(CC)
        .hwndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = lInitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
            CC_FULLOPEN, 0)
    End With

    ' Show the dialog
    If ChooseColor(CC) Then
        'if not canceled, return the color
        ShowColorDialog = CC.rgbResult
        RichTextBox1.BackColor = CC.rgbResult
    Else
        'else return -1
        ShowColorDialog = -1
    End If
End Function









Private Sub Command1_Click()
newborderstyle
End Sub

Private Sub Command2_Click()
Form1.BorderStyle = vbSizable
End Sub

Private Sub Form_GotFocus()
If Form1.WindowState = 1 Then Dialog.Hide
If Form1.Width < 9765 Then Dialog.Show
If Form1.Width > 9765 Then Dialog.Hide
If Form1.Width < 9765 Then Text2.Visible = True
If Form1.Width > 9765 Then Text2.Visible = False

Text3.Text = RichTextBox1.SelFontName
If RichTextBox1.SelAlignment = 0 Then Image24.Visible = True
If Not RichTextBox1.SelAlignment = 0 Then Image24.Visible = False
If Not RichTextBox1.SelAlignment = 0 Then Image10.Visible = True
If RichTextBox1.SelAlignment = 2 Then Image22.Visible = True
If Not RichTextBox1.SelAlignment = 2 Then Image22.Visible = False
If Not RichTextBox1.SelAlignment = 2 Then Image12.Visible = True
If RichTextBox1.SelAlignment = 1 Then Image23.Visible = True
If Not RichTextBox1.SelAlignment = 1 Then Image23.Visible = False
If Not RichTextBox1.SelAlignment = 1 Then Image11.Visible = True
End Sub

Private Sub Form_Load()
Image76.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Visible = True
Label16.Visible = False
Label17.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label43.Visible = False
Label18.Visible = False
Label39.Visible = False
Label42.Visible = False
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = False
Form1.Image52.Visible = False
Form1.Image47.Visible = False
Form1.Image57.Visible = False
Form1.Image50.Visible = False
Form1.Image58.Visible = False
Form1.Image53.Visible = False
Form1.Image59.Visible = False
Form1.Image54.Visible = False
Form1.Image60.Visible = False
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
If Form1.WindowState = 1 Then Dialog.Hide
If Form1.Width < 9765 Then Dialog.Show
If Form1.Width > 9765 Then Dialog.Hide
If Form1.Width < 9765 Then Text2.Visible = True
If Form1.Width > 9765 Then Text2.Visible = False
frmSplash.Show
If Form1.WindowState = 1 Then Dialog.Hide
Text3.Text = RichTextBox1.SelFontName
Text1.Text = RichTextBox1.SelFontSize
If RichTextBox1.SelBold = False Then Image15.Visible = True
If Not RichTextBox1.SelBold = False Then Image16.Visible = True
If RichTextBox1.SelAlignment = 0 Then Image24.Visible = True
If Not RichTextBox1.SelAlignment = 0 Then Image24.Visible = False
If Not RichTextBox1.SelAlignment = 0 Then Image10.Visible = True
If RichTextBox1.SelAlignment = 2 Then Image22.Visible = True
If Not RichTextBox1.SelAlignment = 2 Then Image22.Visible = False
If Not RichTextBox1.SelAlignment = 2 Then Image12.Visible = True
If RichTextBox1.SelAlignment = 1 Then Image23.Visible = True
If Not RichTextBox1.SelAlignment = 1 Then Image23.Visible = False
If Not RichTextBox1.SelAlignment = 1 Then Image11.Visible = True
End Sub

Private Sub Form_LostFocus()
Text3.Text = RichTextBox1.SelFontName
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = RichTextBox1.SelFontName
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line16.Visible = False
Line17.Visible = False
Text3.Text = RichTextBox1.SelFontName
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line13.Visible = False
Line12.Visible = False
Line11.Visible = False
Line14.Visible = False
Line15.Visible = False


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = RichTextBox1.SelFontName
End Sub

Private Sub Form_Resize()
If Form1.Width < 9765 Then Dialog.Show
If Form1.Width > 9765 Then Dialog.Hide
If Form1.Width < 9765 Then Text2.Visible = True
If Form1.Width > 9765 Then Text2.Visible = False
RichTextBox1.Width = Form1.Width
RichTextBox1.Height = Form1.Height
Shape7.Width = Form1.Width
Image41.Width = Form1.Width
Image44.Left = Form1.Width - 460
Shape4.Left = Form1.Width - 495
Shape3.Left = Form1.Width - 495
Image45.Left = Form1.Width - 825
Shape6.Left = Form1.Width - 855
Shape5.Left = Form1.Width - 855
Image46.Left = Form1.Width - 1185
Shape1.Left = Form1.Width - 1215
Shape2.Left = Form1.Width - 1215
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload frmOptions
Unload frmSplash
Unload Dialog
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line14.Visible = True
End Sub

Private Sub Image10_Click()
Image10.Visible = False
Image24.Visible = True
RichTextBox1.SelAlignment = 0
Image23.Visible = False
Image22.Visible = False
Image12.Visible = True
Image11.Visible = True

Dialog.Image10.Visible = False
Dialog.Image24.Visible = True
Dialog.Image23.Visible = False
Dialog.Image22.Visible = False
Dialog.Image12.Visible = True
Dialog.Image11.Visible = True
End Sub

Private Sub Image11_Click()
Image10.Visible = True
Image24.Visible = False
RichTextBox1.SelAlignment = 1
Image23.Visible = True
Image12.Visible = True
Image22.Visible = True
Image11.Visible = False

Dialog.Image10.Visible = True
Dialog.Image24.Visible = False
Dialog.Image23.Visible = True
Dialog.Image12.Visible = True
Dialog.Image22.Visible = True
Dialog.Image11.Visible = False
End Sub

Private Sub Image12_Click()
Image10.Visible = True
Image24.Visible = False
RichTextBox1.SelAlignment = 2
Image23.Visible = False
Image12.Visible = False
Image22.Visible = True
Image11.Visible = True

Dialog.Image10.Visible = True
Dialog.Image24.Visible = False
Dialog.Image23.Visible = False
Dialog.Image12.Visible = False
Dialog.Image22.Visible = True
Dialog.Image11.Visible = True
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
RichTextBox1.SelBold = True
Dialog.Image15.Visible = False
Dialog.Image16.Visible = True
End Sub

Private Sub Image16_Click()
Image15.Visible = True
Image16.Visible = False
RichTextBox1.SelBold = False
Dialog.Image15.Visible = True
Dialog.Image16.Visible = False
End Sub

Private Sub Image17_Click()
Image17.Visible = False
Image18.Visible = True
RichTextBox1.SelItalic = True
Dialog.Image17.Visible = False
Dialog.Image18.Visible = True
End Sub

Private Sub Image18_Click()
Image17.Visible = True
Image18.Visible = False
RichTextBox1.SelItalic = False
Dialog.Image17.Visible = True
Dialog.Image18.Visible = False
End Sub

Private Sub Image19_Click()
Image19.Visible = False
Image20.Visible = True
RichTextBox1.SelUnderline = True
Dialog.Image19.Visible = False
Dialog.Image20.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line15.Visible = True
End Sub

Private Sub Image20_Click()
Image19.Visible = True
Image20.Visible = False
RichTextBox1.SelUnderline = False
Dialog.Image19.Visible = True
Dialog.Image20.Visible = False
End Sub

Private Sub Image21_Click()
Image21.Visible = False
Image26.Visible = True
RichTextBox1.SelBullet = 1
Dialog.Image21.Visible = False
Dialog.Image26.Visible = True
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
RichTextBox1.SelBullet = 0
End Sub

Private Sub Image27_Click()
frmSplash.Show
End Sub

Private Sub Image29_Click()
frmOptions.Show
End Sub

Private Sub Image3_Click()
ShowColorDialog
End Sub


Private Sub Image30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
End Sub

Private Sub Image31_Click()
RichTextBox1.SelText = Clipboard.GetText
End Sub

Private Sub Image31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line11.Visible = True
End Sub

Private Sub Image32_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
End Sub

Private Sub Image32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line12.Visible = True
End Sub

Private Sub Image33_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox1.SelText = ""
End Sub

Private Sub Image33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line13.Visible = True
End Sub

Private Sub Image34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
Line13.Visible = False
Line12.Visible = False
Line11.Visible = False
Line16.Visible = False
Line17.Visible = False
End Sub

Private Sub Image35_Click()
Dialog.Show
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

Private Sub Image38_Click()
FontOpen
End Sub

Private Sub Image38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
Line16.Visible = False
Line17.Visible = False
End Sub

Private Sub Image39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Image4_Click()
ShowColorDialogFore
End Sub


Private Sub Image40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Image41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Image42_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image43_DblClick()
If Form1.WindowState = 2 Then Form1.WindowState = 0
If Not Form1.WindowState = 2 Then Form1.WindowState = 2
End Sub

Private Sub Image43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image44_Click()
Unload Form1
Unload Dialog
Unload frmOptions
Unload frmSplash
End Sub

Private Sub Image45_Click()
Form1.WindowState = 2
End Sub

Private Sub Image46_Click()
Form1.WindowState = 1
End Sub

Private Sub Image47_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image48_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image49_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image5_Click()
RichTextBox1.Text = ""
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = True
Line9.Visible = False
Line10.Visible = False
End Sub

Private Sub Image50_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image50_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image50_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image52_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image52_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image52_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image53_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image53_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image53_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image54_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image54_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image55_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image55_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image55_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image56_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image56_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image57_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image57_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image58_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image58_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image58_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image59_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image59_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image59_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image60_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image60_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image60_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image61_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image61_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image62_DblClick()
If Form1.WindowState = 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 Then Form1.WindowState = 0
End Sub

Private Sub Image62_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Image62_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub Image63_DblClick()
Unload Form1
Unload Dialog
Unload frmOptions
Unload frmSplash
End Sub

Private Sub Image63_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Image64_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Text = ""
End Sub

Private Sub Image65_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
OpenFile
End Sub

Private Sub Image66_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
salva_bxy
End Sub

Private Sub Image67_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
frmSplash.Show
End Sub

Private Sub Image68_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox1.SelText = ""

Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False

End Sub

Private Sub Image69_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
End Sub

Private Sub Image7_Click()
FontOpen
End Sub

Private Sub Image70_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
RichTextBox1.SelText = Clipboard.GetText
End Sub

Private Sub Image71_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
SendKeys "^Z"
End Sub

Private Sub Image72_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
SendKeys "^Y"
End Sub

Private Sub Image73_Click()
AddFile
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub

Private Sub Image74_Click()
Dialog.Show
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub

Private Sub Image75_Click()
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Image21.Visible = False
Image26.Visible = True
RichTextBox1.SelBullet = 1
Dialog.Image21.Visible = False
Dialog.Image26.Visible = True
End Sub

Private Sub Image76_Click()
frmOptions.Show
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub

Private Sub Image78_Click()
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Image10.Visible = True
Image24.Visible = False
RichTextBox1.SelAlignment = 1
Image23.Visible = True
Image12.Visible = True
Image22.Visible = True
Image11.Visible = False

Dialog.Image10.Visible = True
Dialog.Image24.Visible = False
Dialog.Image23.Visible = True
Dialog.Image12.Visible = True
Dialog.Image22.Visible = True
Dialog.Image11.Visible = False
End Sub

Private Sub Image79_Click()
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Image10.Visible = True
Image24.Visible = False
RichTextBox1.SelAlignment = 2
Image23.Visible = False
Image12.Visible = False
Image22.Visible = True
Image11.Visible = True

Dialog.Image10.Visible = True
Dialog.Image24.Visible = False
Dialog.Image23.Visible = False
Dialog.Image12.Visible = False
Dialog.Image22.Visible = True
Dialog.Image11.Visible = True
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Image80_Click()
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Image10.Visible = False
Image24.Visible = True
RichTextBox1.SelAlignment = 0
Image23.Visible = False
Image22.Visible = False
Image12.Visible = True
Image11.Visible = True

Dialog.Image10.Visible = False
Dialog.Image24.Visible = True
Dialog.Image23.Visible = False
Dialog.Image22.Visible = False
Dialog.Image12.Visible = True
Dialog.Image11.Visible = True
End Sub

Private Sub Image81_Click()
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
ShowColorDialogFore
End Sub

Private Sub Image82_Click()
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
ShowColorDialog
End Sub

Private Sub Image9_Click()
frmOptions.Show
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image29.Visible = True
End Sub


Private Sub Label11_Click()
FontOpen
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.Font.Italic = True
Label4.Font.Italic = False
Label8.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Label12_Click()
Text3.Visible = False
Text1.Visible = False
Image76.Visible = False
Label12.Visible = False
Label19.Visible = True
Shape15.Visible = True
Shape14.Visible = True
RichTextBox1.Visible = False
Image78.Visible = True
Image79.Visible = True
Image80.Visible = True
Label32.Visible = True
Label33.Visible = True
Label34.Visible = True
Label17.Visible = False
Label8.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.Font.Italic = True
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Label13_Click()
Text3.Visible = False
Text1.Visible = False
RichTextBox1.Visible = False
Image72.Visible = False
Image76.Visible = True
Label17.Visible = False
Label8.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = False
Label20.Visible = True
Shape13.Visible = True
Shape12.Visible = True
Image73.Visible = True
Image74.Visible = True
Image75.Visible = True
Label35.Visible = True
Label36.Visible = True
Label37.Visible = True
Label38.Visible = True
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.Font.Italic = True
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Label14_Click()
Text3.Visible = False
Text1.Visible = False
Image76.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label17.Visible = False
Label8.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label14.Visible = False
Label21.Visible = True
Shape17.Visible = True
Shape16.Visible = True
Image81.Visible = True
Image82.Visible = True
Label40.Visible = True
Label41.Visible = True
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.Font.Italic = True
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Label15_Click()
Form1.Image43.Visible = True
Form1.Image42.Visible = True
newborderstyle
Shape7.Visible = False
Image63.Visible = False
Label4.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label4.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False

Image76.Visible = False

Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False

Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False

Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False

Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label17.Visible = False

RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False

Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False

Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.Font.Italic = True
Label4.Font.Italic = False
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
End Sub

Private Sub Label16_Click()
Image76.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Visible = True
End Sub

Private Sub Label17_Click()
Image76.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Visible = True
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label42.Font.Italic = False
Label39.Font.Italic = False
Label18.Font.Italic = True
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox1.SelText = ""

Label22.Visible = False
Label39.Visible = False
Label42.Visible = False
Label43.Visible = False
End Sub

Private Sub Label19_Click()
Text3.Visible = True
Text1.Visible = True
Image76.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Visible = True
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line14.Visible = False
End Sub

Private Sub Label20_Click()
Text3.Visible = True
Text1.Visible = True
Image76.Visible = False
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Visible = True
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False

End Sub

Private Sub Label21_Click()
Text3.Visible = True
Text1.Visible = True
Image76.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Visible = True
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
End Sub

Private Sub Label22_Click()
Label22.Visible = False
Label39.Visible = False
Label42.Visible = False
Label43.Visible = False
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label42.Font.Italic = False
Label39.Font.Italic = False
Label18.Font.Italic = False
End Sub

Private Sub Label23_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
RichTextBox1.Text = ""
End Sub

Private Sub Label24_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
OpenFile
End Sub

Private Sub Label25_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
salva_bxy
End Sub

Private Sub Label26_Click()
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
frmSplash.Show
End Sub

Private Sub Label27_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
End Sub

Private Sub Label28_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
End Sub

Private Sub Label29_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
End Sub

Private Sub Label3_DblClick()
Unload Form1
Unload Dialog
Unload frmOptions
Unload frmSplash
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line15.Visible = False
End Sub

Private Sub Label30_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
End Sub

Private Sub Label31_Click()
Label17.Visible = False
Label8.Visible = True
RichTextBox1.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
End Sub

Private Sub Label32_Click()
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Image10.Visible = True
Image24.Visible = False
RichTextBox1.SelAlignment = 1
Image23.Visible = True
Image12.Visible = True
Image22.Visible = True
Image11.Visible = False

Dialog.Image10.Visible = True
Dialog.Image24.Visible = False
Dialog.Image23.Visible = True
Dialog.Image12.Visible = True
Dialog.Image22.Visible = True
Dialog.Image11.Visible = False
End Sub

Private Sub Label33_Click()
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Image10.Visible = True
Image24.Visible = False
RichTextBox1.SelAlignment = 2
Image23.Visible = False
Image12.Visible = False
Image22.Visible = True
Image11.Visible = True

Dialog.Image10.Visible = True
Dialog.Image24.Visible = False
Dialog.Image23.Visible = False
Dialog.Image12.Visible = False
Dialog.Image22.Visible = True
Dialog.Image11.Visible = True
End Sub

Private Sub Label34_Click()
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
RichTextBox1.Visible = True
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Image10.Visible = False
Image24.Visible = True
RichTextBox1.SelAlignment = 0
Image23.Visible = False
Image22.Visible = False
Image12.Visible = True
Image11.Visible = True

Dialog.Image10.Visible = False
Dialog.Image24.Visible = True
Dialog.Image23.Visible = False
Dialog.Image22.Visible = False
Dialog.Image12.Visible = True
Dialog.Image11.Visible = True
End Sub

Private Sub Label35_Click()
AddFile
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub

Private Sub Label36_Click()
Dialog.Show
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub

Private Sub Label37_Click()
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Image21.Visible = False
Image26.Visible = True
RichTextBox1.SelBullet = 1
Dialog.Image21.Visible = False
Dialog.Image26.Visible = True
End Sub

Private Sub Label38_Click()
frmOptions.Show
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
End Sub

Private Sub Label39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
Label42.Font.Italic = False
Label39.Font.Italic = True
Label18.Font.Italic = False
Label22.Visible = False
Label39.Visible = False
Label42.Visible = False
Label43.Visible = False
End Sub

Private Sub Label4_Click()
Image76.Visible = False
Label4.Visible = False
Label16.Visible = True
Shape9.Visible = True
Shape8.Visible = True
Image64.Visible = True
Image65.Visible = True
Image66.Visible = True
Image67.Visible = True
Line1.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
RichTextBox1.Visible = False
Label17.Visible = False
Label8.Visible = True
Shape11.Visible = False
Shape10.Visible = False
Image68.Visible = False
Image69.Visible = False
Image70.Visible = False
Image71.Visible = False
Image72.Visible = False
Line2.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image72.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Font.Italic = True
Label8.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Label40_Click()
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
ShowColorDialogFore
End Sub

Private Sub Label41_Click()
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
ShowColorDialog
End Sub

Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label42.Font.Italic = True
Label39.Font.Italic = False
Label18.Font.Italic = False
RichTextBox1.SelText = Clipboard.GetText
Label22.Visible = False
Label39.Visible = False
Label42.Visible = False
Label43.Visible = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = True
Image29.Visible = False
End Sub

Private Sub Label8_Click()
Image72.Visible = True
Image76.Visible = False
Label17.Visible = True
Label8.Visible = False
RichTextBox1.Visible = False
Shape11.Visible = True
Shape10.Visible = True
Image68.Visible = True
Image69.Visible = True
Image70.Visible = True
Image71.Visible = True
Image72.Visible = True
Line2.Visible = True
Label27.Visible = True
Label28.Visible = True
Label29.Visible = True
Label30.Visible = True
Label31.Visible = True
Label14.Visible = True
Label21.Visible = False
Shape17.Visible = False
Shape16.Visible = False
Image81.Visible = False
Image82.Visible = False
Label40.Visible = False
Label41.Visible = False
Label13.Visible = True
Label20.Visible = False
Shape13.Visible = False
Shape12.Visible = False
Image73.Visible = False
Image74.Visible = False
Image75.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label4.Visible = True
Label16.Visible = False
Shape9.Visible = False
Shape8.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
Line1.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label12.Visible = True
Label19.Visible = False
Shape15.Visible = False
Shape14.Visible = False
Image78.Visible = False
Image79.Visible = False
Image80.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.Font.Italic = True
Label4.Font.Italic = False
Label11.Font.Italic = False
Label12.Font.Italic = False
Label13.Font.Italic = False
Label14.Font.Italic = False
Label15.Font.Italic = False
End Sub

Private Sub Label9_Click()
AddFile
End Sub

Private Sub Newstyle_Click(Index As Integer)
newborderstyle
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Label22.Left = mlngX
Label22.Top = mlngY
Label39.Left = mlngX + 660
Label39.Top = mlngY + 165
Label18.Visible = True
Label18.Left = mlngX + 140
Label18.Top = mlngY + 165
Label42.Left = mlngX + 1260
Label42.Top = mlngY + 165
Label43.Left = mlngX + 60
Label43.Top = mlngY + 45
Label22.Visible = True
Label39.Visible = True
Label42.Visible = True
Label43.Visible = True

    End If
    
If Button = 1 Then
Label43.Visible = False
Label22.Visible = False
Label18.Visible = False
Label23.Visible = False
Label39.Visible = False
Label42.Visible = False
End If
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line16.Visible = False
Line17.Visible = False
Text3.Text = RichTextBox1.SelFontName
Dialog.Text3.Text = RichTextBox1.SelFontName
Dialog.Text1.Text = RichTextBox1.SelFontSize
Image9.Visible = True
Image29.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line13.Visible = False
Line12.Visible = False
Line11.Visible = False
Line14.Visible = False
Line15.Visible = False
End Sub

Private Sub Text1_Change()
RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub Text1_GotFocus()
RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub Text1_LostFocus()
RichTextBox1.SelFontSize = Text1.Text
End Sub

Private Sub RichTextBox1_Change()
Text3.Text = RichTextBox1.SelFontName
Text1.Text = RichTextBox1.SelFontSize
Dialog.Text3.Text = RichTextBox1.SelFontName
Dialog.Text1.Text = RichTextBox1.SelFontSize


End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line15.Visible = True
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
RichTextBox1.SelFontName = Text3.Text
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = RichTextBox1.SelFontName
Line14.Visible = True
End Sub

