VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Starboy Notepad"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   17880
   Icon            =   "NotepadForm1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   17880
   StartUpPosition =   3  'Windows Default
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
      Height          =   1215
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "The controls in this window have been moved to the Tools window. Once you resize, they'll get restored."
      Top             =   0
      Width           =   15255
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lines:"
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
      Left            =   3720
      TabIndex        =   78
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Word count"
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
      Left            =   1800
      TabIndex        =   77
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Word count"
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
      Left            =   0
      TabIndex        =   76
      Top             =   6600
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   555
      Width           =   495
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   11040
      TabIndex        =   61
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   240
         TabIndex        =   68
         Top             =   1680
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthBackColor  =   16777215
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   80478210
         TitleBackColor  =   15311702
         TitleForeColor  =   16777215
         TrailingForeColor=   12632256
         CurrentDate     =   44415
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Add date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
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
         Left            =   2400
         TabIndex        =   71
         Top             =   1400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   1560
         TabIndex        =   70
         Top             =   1400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         Left            =   840
         TabIndex        =   69
         Top             =   1400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   65
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   64
         Text            =   "New task..."
         Top             =   900
         Width           =   2535
      End
      Begin VB.Line Line28 
         Visible         =   0   'False
         X1              =   1320
         X2              =   1440
         Y1              =   1560
         Y2              =   1440
      End
      Begin VB.Line Line29 
         Visible         =   0   'False
         X1              =   2160
         X2              =   2280
         Y1              =   1560
         Y2              =   1440
      End
      Begin VB.Image Image91 
         Height          =   375
         Left            =   720
         Picture         =   "NotepadForm1.frx":06EA
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label87 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   67
         ToolTipText     =   "Click here to select the date from a calendar."
         Top             =   1400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label86 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   66
         Top             =   1800
         Width           =   975
      End
      Begin VB.Image Image90 
         Height          =   255
         Left            =   2760
         Picture         =   "NotepadForm1.frx":6592C
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image Image89 
         Height          =   375
         Left            =   240
         Picture         =   "NotepadForm1.frx":65EB6
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label85 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   255
         Left            =   2880
         TabIndex        =   63
         Top             =   360
         Width           =   255
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   3000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   11520
      TabIndex        =   54
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label Label83 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   60
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label82 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copy path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   59
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label81 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   240
         TabIndex        =   58
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label80 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   3120
         Width           =   975
      End
      Begin VB.Line Line26 
         BorderColor     =   &H80000000&
         X1              =   240
         X2              =   2520
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label79 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   1935
         Left            =   240
         TabIndex        =   56
         Top             =   840
         Width           =   2295
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   2520
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "File info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   14640
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Characters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Word count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Line count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   79
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Line count bar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ribbon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   37
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Activate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   75
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "By activating this option, you'll activate a new mode in Starboy, which substitutes the default Windows one with a custom one."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   74
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   2880
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "New Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label70 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   360
         Width           =   255
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   2880
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label69 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit mode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   "Hide or show"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   15000
      TabIndex        =   44
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   50
         Top             =   3120
         Width           =   2295
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   1335
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2355
         _Version        =   393217
         ScrollBars      =   3
         MousePointer    =   3
         Appearance      =   0
         TextRTF         =   $"NotepadForm1.frx":CB0F8
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   255
         Left            =   2400
         TabIndex        =   52
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         ToolTipText     =   "Restore the selected record in the clipboard list. Note that this won't be copying the formatting."
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   49
         ToolTipText     =   "Save the text currently in the clipboard in the clipboard list. Note that formatting won't be saved with this option."
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         ToolTipText     =   "If the space available here is too small, you can click here and open a bigger clipboard."
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "Show here in this sidebar all the contents of the clipboard."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   46
         Top             =   720
         Width           =   2295
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00C0C000&
         X1              =   240
         X2              =   2520
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Clipboard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   5295
      Left            =   0
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17640
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   4080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   210
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   14280
      TabIndex        =   21
      ToolTipText     =   "Typewriter [Palatino Linotype, 12, Black]"
      Top             =   120
      Width           =   735
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   495
         TabIndex        =   29
         ToolTipText     =   "Typewriter [Palatino Linotype, 12, Black]"
         Top             =   375
         Width           =   105
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   75
         TabIndex        =   25
         ToolTipText     =   "Typewriter [Palatino Linotype, 12, Black]"
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13440
      TabIndex        =   20
      ToolTipText     =   "Subtitle [Tahoma, 9, Grey]"
      Top             =   120
      Width           =   735
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   480
         TabIndex        =   28
         ToolTipText     =   "Subtitle [Tahoma, 9, Grey]"
         Top             =   405
         Width           =   120
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   630
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Subtitle [Tahoma, 9, Grey]"
         Top             =   105
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12600
      TabIndex        =   19
      ToolTipText     =   "Title [Arial, 28, Bold, Blue]"
      Top             =   120
      Width           =   735
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E9A356&
         Height          =   255
         Left            =   495
         TabIndex        =   27
         ToolTipText     =   "Title [Arial, 28, Bold, Blue]"
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label55 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E9A356&
         Height          =   630
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Title [Arial, 28, Bold, Blue]"
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11760
      TabIndex        =   18
      ToolTipText     =   "Normal [Tahoma, 12, Black]"
      Top             =   120
      Width           =   735
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   26
         ToolTipText     =   "Normal [Tahoma, 12, Black]"
         Top             =   405
         Width           =   120
      End
      Begin VB.Label Label54 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Normal [Tahoma, 12, Black]"
         Top             =   75
         Width           =   330
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "View Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5415
      Left            =   0
      TabIndex        =   83
      Top             =   1200
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   9551
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"NotepadForm1.frx":CB17A
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9480
      TabIndex        =   84
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   0
      X2              =   18000
      Y1              =   1400
      Y2              =   1400
   End
   Begin VB.Image Image92 
      Height          =   330
      Left            =   17280
      Picture         =   "NotepadForm1.frx":CB205
      Stretch         =   -1  'True
      ToolTipText     =   "A quick and easy to use to-do list, built directly in Starboy"
      Top             =   675
      Width           =   435
   End
   Begin VB.Image Image88 
      Height          =   195
      Left            =   1200
      Picture         =   "NotepadForm1.frx":DA097
      Stretch         =   -1  'True
      ToolTipText     =   "Shows info about the current opened file"
      Top             =   210
      Width           =   195
   End
   Begin VB.Image Image87 
      Height          =   195
      Left            =   1200
      Picture         =   "NotepadForm1.frx":DA781
      Stretch         =   -1  'True
      ToolTipText     =   "(BETA) Prints the current file"
      Top             =   555
      Width           =   195
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1200
      X2              =   1410
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   1200
      X2              =   1410
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label77 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   3120
      TabIndex        =   53
      Top             =   240
      Width           =   90
   End
   Begin VB.Image Image85 
      Height          =   330
      Left            =   16800
      Picture         =   "NotepadForm1.frx":EA43F
      Stretch         =   -1  'True
      ToolTipText     =   "Find and replace the text you want"
      Top             =   675
      Width           =   375
   End
   Begin VB.Label Label68 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   17640
      TabIndex        =   40
      Top             =   945
      Width           =   90
   End
   Begin VB.Label Label66 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   17640
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label65 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   31
      Top             =   1200
      Width           =   135
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   14280
      X2              =   15000
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   13440
      X2              =   14160
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   12600
      X2              =   13320
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C000&
      Visible         =   0   'False
      X1              =   11760
      X2              =   12480
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00E9A356&
      X1              =   11760
      X2              =   12480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   15360
      X2              =   15360
      Y1              =   1080
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   11400
      X2              =   11400
      Y1              =   1080
      Y2              =   120
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "FILE"
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
      Left            =   615
      TabIndex        =   16
      Top             =   885
      Width           =   345
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "TEXT ACTIONS"
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
      Left            =   1920
      TabIndex        =   15
      Top             =   885
      Width           =   1095
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "FONT"
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
      Left            =   4350
      TabIndex        =   14
      Top             =   885
      Width           =   435
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "ALIGNMENT"
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
      Left            =   6120
      TabIndex        =   13
      Top             =   885
      Width           =   885
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "TEXT TOOLS"
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
      Left            =   7890
      TabIndex        =   12
      Top             =   885
      Width           =   915
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOOLS"
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
      Left            =   9465
      TabIndex        =   11
      Top             =   885
      Width           =   1515
   End
   Begin VB.Image Image77 
      Height          =   330
      Left            =   16800
      Picture         =   "NotepadForm1.frx":F6889
      Stretch         =   -1  'True
      ToolTipText     =   "(BETA) Checks if the text you have written is correctly written or not."
      Top             =   240
      Width           =   915
   End
   Begin VB.Image Image83 
      Height          =   330
      Left            =   16800
      Picture         =   "NotepadForm1.frx":10F4D7
      Stretch         =   -1  'True
      ToolTipText     =   "Checks if the text you have written is correctly written or not."
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label46 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   16680
      TabIndex        =   10
      Top             =   165
      Width           =   1455
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
      Left            =   7920
      TabIndex        =   5
      Top             =   225
      Width           =   855
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00C0C000&
      X1              =   2880
      X2              =   3090
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C0C000&
      X1              =   1920
      X2              =   2280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image36 
      Height          =   435
      Left            =   1860
      Picture         =   "NotepadForm1.frx":127F49
      Stretch         =   -1  'True
      ToolTipText     =   "Undo (CTRL+Z)"
      Top             =   270
      Width           =   435
   End
   Begin VB.Image Image35 
      Height          =   330
      Left            =   16200
      Picture         =   "NotepadForm1.frx":156A07
      Stretch         =   -1  'True
      ToolTipText     =   "Opens the external Tools window"
      Top             =   675
      Width           =   495
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0C000&
      X1              =   4080
      X2              =   4560
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C000&
      X1              =   4080
      X2              =   5400
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0C000&
      X1              =   2505
      X2              =   2715
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C000&
      X1              =   2520
      X2              =   2730
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C000&
      X1              =   2880
      X2              =   3090
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C000&
      X1              =   840
      X2              =   1050
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C000&
      X1              =   840
      X2              =   1050
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C000&
      X1              =   240
      X2              =   720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image25 
      Height          =   195
      Left            =   840
      Picture         =   "NotepadForm1.frx":17C265
      Stretch         =   -1  'True
      ToolTipText     =   "Saves the current file"
      Top             =   555
      Width           =   195
   End
   Begin VB.Image Image13 
      Height          =   195
      Left            =   840
      Picture         =   "NotepadForm1.frx":17C94F
      Stretch         =   -1  'True
      ToolTipText     =   "Opens a new file"
      Top             =   210
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   435
      Left            =   255
      Picture         =   "NotepadForm1.frx":17D039
      Stretch         =   -1  'True
      ToolTipText     =   "Cleans the existing file. Note: all unsaved changes will be discarded"
      Top             =   270
      Width           =   435
   End
   Begin VB.Image Image30 
      Height          =   1080
      Left            =   120
      Picture         =   "NotepadForm1.frx":17D723
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1365
   End
   Begin VB.Image Image9 
      Height          =   330
      Left            =   15720
      Picture         =   "NotepadForm1.frx":206761
      Stretch         =   -1  'True
      ToolTipText     =   "Settings"
      Top             =   675
      Width           =   345
   End
   Begin VB.Image Image27 
      Height          =   330
      Left            =   15600
      Picture         =   "NotepadForm1.frx":216BCB
      Stretch         =   -1  'True
      ToolTipText     =   "Returns to the splash screen"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image29 
      Height          =   330
      Left            =   15720
      Picture         =   "NotepadForm1.frx":260515
      Stretch         =   -1  'True
      ToolTipText     =   "Settings"
      Top             =   675
      Width           =   345
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15600
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image21 
      Height          =   210
      Left            =   7680
      Picture         =   "NotepadForm1.frx":26FB1B
      Stretch         =   -1  'True
      ToolTipText     =   "Adds text bullets to your current row"
      Top             =   210
      Width           =   255
   End
   Begin VB.Image Image26 
      Height          =   210
      Left            =   7680
      Picture         =   "NotepadForm1.frx":2843B5
      Stretch         =   -1  'True
      ToolTipText     =   "Removes text bullets from your current row"
      Top             =   210
      Width           =   255
   End
   Begin VB.Image Image12 
      Height          =   345
      Left            =   6360
      Picture         =   "NotepadForm1.frx":298D57
      Stretch         =   -1  'True
      ToolTipText     =   "Center alignment"
      Top             =   315
      Width           =   345
   End
   Begin VB.Image Image11 
      Height          =   345
      Left            =   6840
      Picture         =   "NotepadForm1.frx":2A5FF9
      Stretch         =   -1  'True
      ToolTipText     =   "Right alignment"
      Top             =   315
      Width           =   345
   End
   Begin VB.Image Image10 
      Height          =   345
      Left            =   5880
      Picture         =   "NotepadForm1.frx":2B3427
      Stretch         =   -1  'True
      ToolTipText     =   "Left alignment"
      Top             =   315
      Width           =   345
   End
   Begin VB.Image Image24 
      Height          =   345
      Left            =   5880
      Picture         =   "NotepadForm1.frx":2C0855
      Stretch         =   -1  'True
      ToolTipText     =   "Left alignment"
      Top             =   315
      Width           =   345
   End
   Begin VB.Image Image23 
      Height          =   345
      Left            =   6840
      Picture         =   "NotepadForm1.frx":2CD9F7
      Stretch         =   -1  'True
      ToolTipText     =   "Right alignment"
      Top             =   315
      Width           =   345
   End
   Begin VB.Image Image22 
      Height          =   345
      Left            =   6360
      Picture         =   "NotepadForm1.frx":2DA879
      Stretch         =   -1  'True
      ToolTipText     =   "Center alignment"
      Top             =   315
      Width           =   345
   End
   Begin VB.Image Image19 
      Height          =   270
      Left            =   5160
      Picture         =   "NotepadForm1.frx":2E7DC7
      Stretch         =   -1  'True
      ToolTipText     =   "Underline"
      Top             =   555
      Width           =   270
   End
   Begin VB.Image Image17 
      Height          =   270
      Left            =   4920
      Picture         =   "NotepadForm1.frx":2F4E49
      Stretch         =   -1  'True
      ToolTipText     =   "Italic"
      Top             =   555
      Width           =   270
   End
   Begin VB.Image Image15 
      Height          =   270
      Left            =   4680
      Picture         =   "NotepadForm1.frx":302A7B
      Stretch         =   -1  'True
      ToolTipText     =   "Bold"
      Top             =   555
      Width           =   270
   End
   Begin VB.Image Image20 
      Height          =   270
      Left            =   5160
      Picture         =   "NotepadForm1.frx":3101E5
      Stretch         =   -1  'True
      ToolTipText     =   "Underline"
      Top             =   555
      Width           =   270
   End
   Begin VB.Image Image18 
      Height          =   270
      Left            =   4920
      Picture         =   "NotepadForm1.frx":31D59F
      Stretch         =   -1  'True
      ToolTipText     =   "Italic"
      Top             =   555
      Width           =   270
   End
   Begin VB.Image Image16 
      Height          =   270
      Left            =   4680
      Picture         =   "NotepadForm1.frx":32A959
      Stretch         =   -1  'True
      ToolTipText     =   "Bold"
      Top             =   555
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
      Left            =   7920
      TabIndex        =   1
      Top             =   555
      Width           =   1095
   End
   Begin VB.Image Image14 
      Height          =   225
      Left            =   7695
      Picture         =   "NotepadForm1.frx":338143
      Stretch         =   -1  'True
      ToolTipText     =   "Adds text to this file from another file"
      Top             =   555
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
      Left            =   10080
      TabIndex        =   0
      Top             =   570
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   10560
      Picture         =   "NotepadForm1.frx":34DD85
      Stretch         =   -1  'True
      ToolTipText     =   "Changes the colour of the selected font"
      Top             =   570
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   10560
      Picture         =   "NotepadForm1.frx":34E46F
      Stretch         =   -1  'True
      ToolTipText     =   "Changes the colour of the background"
      Top             =   240
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   3960
      Picture         =   "NotepadForm1.frx":34EB59
      Stretch         =   -1  'True
      Top             =   525
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3960
      Picture         =   "NotepadForm1.frx":3B3D9B
      Stretch         =   -1  'True
      Top             =   165
      Width           =   1485
   End
   Begin VB.Image Image39 
      Height          =   1080
      Left            =   5760
      Picture         =   "NotepadForm1.frx":418FDD
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1575
   End
   Begin VB.Image Image8 
      Height          =   1080
      Left            =   7560
      Picture         =   "NotepadForm1.frx":4BDF5F
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1575
   End
   Begin VB.Image Image40 
      Height          =   1080
      Left            =   9360
      Picture         =   "NotepadForm1.frx":5693B5
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image Image31 
      Height          =   195
      Left            =   2880
      Picture         =   "NotepadForm1.frx":625A4F
      Stretch         =   -1  'True
      ToolTipText     =   "Paste (CTRL+V)"
      Top             =   210
      Width           =   180
   End
   Begin VB.Image Image33 
      Height          =   195
      Left            =   2520
      Picture         =   "NotepadForm1.frx":626139
      Stretch         =   -1  'True
      ToolTipText     =   "Cut (CTRL+Z)"
      Top             =   210
      Width           =   195
   End
   Begin VB.Image Image32 
      Height          =   195
      Left            =   2520
      Picture         =   "NotepadForm1.frx":626823
      Stretch         =   -1  'True
      ToolTipText     =   "Copy (CTRL+C)"
      Top             =   555
      Width           =   195
   End
   Begin VB.Image Image37 
      Height          =   195
      Left            =   2880
      Picture         =   "NotepadForm1.frx":626F0D
      Stretch         =   -1  'True
      ToolTipText     =   "Redo (CTRL+Y)"
      Top             =   555
      Width           =   180
   End
   Begin VB.Image Image34 
      Height          =   1080
      Left            =   1680
      Picture         =   "NotepadForm1.frx":6563BF
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1635
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   3600
      Picture         =   "NotepadForm1.frx":70F201
      Stretch         =   -1  'True
      Top             =   555
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   225
      Left            =   3600
      Picture         =   "NotepadForm1.frx":72E08F
      Stretch         =   -1  'True
      ToolTipText     =   "Change the font of your text"
      Top             =   210
      Width           =   255
   End
   Begin VB.Image Image38 
      Height          =   1080
      Left            =   3540
      Picture         =   "NotepadForm1.frx":72E779
      Stretch         =   -1  'True
      ToolTipText     =   "Change the font of your text"
      Top             =   75
      Width           =   2055
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "STYLES"
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
      Left            =   13110
      TabIndex        =   17
      Top             =   885
      Width           =   555
   End
   Begin VB.Image Image84 
      Height          =   1080
      Left            =   11400
      Picture         =   "NotepadForm1.frx":80D8CB
      Stretch         =   -1  'True
      Top             =   75
      Width           =   3975
   End
   Begin VB.Image Image41 
      Height          =   1200
      Left            =   0
      Picture         =   "NotepadForm1.frx":8A4F91
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17955
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   30
      ToolTipText     =   "Line count (total)"
      Top             =   1560
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   12600
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Image Image63 
      Height          =   255
      Left            =   120
      Picture         =   "NotepadForm1.frx":A56903
      Stretch         =   -1  'True
      Top             =   45
      Width           =   255
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   0
      Width           =   10575
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
      TabIndex        =   6
      Top             =   75
      Width           =   13815
   End
   Begin VB.Image Image46 
      Height          =   210
      Left            =   11190
      Picture         =   "NotepadForm1.frx":A56FED
      Stretch         =   -1  'True
      Top             =   70
      Width           =   210
   End
   Begin VB.Image Image45 
      Height          =   220
      Left            =   11550
      Picture         =   "NotepadForm1.frx":A5EE0F
      Stretch         =   -1  'True
      Top             =   70
      Width           =   220
   End
   Begin VB.Image Image44 
      Height          =   220
      Left            =   11915
      Picture         =   "NotepadForm1.frx":A67EB1
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
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image43 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A703D7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image42 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A74811
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image47 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A78C4B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image57 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A7E85D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image56 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A8446F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image55 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A87D85
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image54 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A8DF79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image53 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A9409C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image50 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A99911
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image48 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":A9F99D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image49 
      Height          =   360
      Left            =   -120
      Picture         =   "NotepadForm1.frx":AA4F2E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image51 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":AAA4BF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12315
   End
   Begin VB.Image Image52 
      Height          =   360
      Left            =   0
      Picture         =   "NotepadForm1.frx":AB0168
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image60 
      Height          =   360
      Left            =   -240
      Picture         =   "NotepadForm1.frx":AB5E11
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image59 
      Height          =   360
      Left            =   -360
      Picture         =   "NotepadForm1.frx":ABBF34
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image58 
      Height          =   360
      Left            =   -360
      Picture         =   "NotepadForm1.frx":AC17A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image62 
      Height          =   360
      Left            =   -120
      Picture         =   "NotepadForm1.frx":AC7835
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Image Image61 
      Height          =   360
      Left            =   -120
      Picture         =   "NotepadForm1.frx":ACB14B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   38835
   End
   Begin VB.Menu mnuPUForm 
      Caption         =   "Form Popup"
      Visible         =   0   'False
      Begin VB.Menu undo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu redo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu clipboardfuck 
         Caption         =   "Clipboard"
      End
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu spell 
         Caption         =   "Spell Check"
      End
      Begin VB.Menu find 
         Caption         =   "Find"
      End
      Begin VB.Menu tasks 
         Caption         =   "Tasks"
      End
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










Private Sub Command2_Click()
Form1.BorderStyle = vbSizable
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then Label87.Visible = True
If Check1.Value = 0 Then Label87.Visible = False
If Check1.Value = 1 Then Line28.Visible = True
If Check1.Value = 0 Then Line28.Visible = False
If Check1.Value = 1 Then Line29.Visible = True
If Check1.Value = 0 Then Line29.Visible = False
If Check1.Value = 1 Then Image91.Visible = True
If Check1.Value = 0 Then Image91.Visible = False
If Check1.Value = 1 Then Text6.Visible = True
If Check1.Value = 0 Then Text6.Visible = False
If Check1.Value = 1 Then Text7.Visible = True
If Check1.Value = 0 Then Text7.Visible = False
If Check1.Value = 1 Then Text8.Visible = True
If Check1.Value = 0 Then Text8.Visible = False

End Sub

Private Sub Check2_Click()


If Check2.Value = 1 & Image84.Top = 75 Then RichTextBox1.Top = 1200
If Check2.Value = 1 & Image84.Top = 315 Then RichTextBox1.Top = 1440
If Check2.Value = 0 & Image84.Top = 75 Then RichTextBox1.Top = 0
If Check2.Value = 0 & Image84.Top = 315 Then RichTextBox1.Top = 360

'1440, 360

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then List1.Visible = True
If Check3.Value = 0 Then List1.Visible = False
If Check3.Value = 0 Then RichTextBox1.Left = 240
If Check3.Value = 0 Then List1.Width = 255
If Check3.Value = 1 Then List1.Width = 375
If Check3.Value = 1 Then RichTextBox1.Left = 360
End Sub

Private Sub Command1_Click()
Dim i As Integer
For i = "1" To RichTextBox2.Text
Form1.List2.AddItem (i)
Next
End Sub

Private Sub Combo1_Click()
MonthView1.Visible = True
End Sub

Private Sub Combo1_DropDown()
If MonthView1.Visible = True Then MonthView1.Visible = False
If MonthView1.Visible = False Then MonthView1.Visible = True
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then Frame12.Visible = True
If Check4.Value = 0 Then Frame12.Visible = False
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then Frame11.Visible = True
If Check5.Value = 0 Then Frame11.Visible = False
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then Frame10.Visible = True
If Check6.Value = 0 Then Frame10.Visible = False
End Sub

Private Sub clipboard_Click()
Frame7.Visible = True
End Sub

Private Sub Check7_Click()

End Sub

Private Sub copy_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox2.Text = ""
RichTextBox2.TextRTF = RichTextBox1.SelRTF
End Sub

Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox2.Text = ""
RichTextBox2.TextRTF = RichTextBox1.SelRTF
RichTextBox1.SelRTF = ""
End Sub

Private Sub find_Click()
SendKeys "^A"
frmFind.Show
End Sub

Private Sub Form_GotFocus()

If Form1.WindowState = 1 Then Dialog.Hide
If Form1.WindowState = 1 Then Dialog.Hide
If Form1.Width < 9765 And Form1.WindowState = 0 Then Dialog.Show
If Form1.Width > 9765 And Form1.WindowState = 1 Then Dialog.Hide
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
RichTextBox1.Top = 1200
If Form1.WindowState = 1 Then Dialog.Hide
If Form1.WindowState = 2 Then Label44.Visible = True
If Not Form1.WindowState = 2 Then Label44.Visible = False

RichTextBox1.Visible = True



RichTextBox1.Visible = True











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
If Form1.WindowState = 1 Then Dialog.Hide
Text3.Text = RichTextBox1.SelFontName
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = RichTextBox1.SelFontName

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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = RichTextBox1.SelFontName

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form1.WindowState = 2 Then Label44.Visible = True
If Not Form1.WindowState = 2 Then Label44.Visible = False
If Form1.WindowState = 1 Then Dialog.Hide
If Form1.Width < 9765 And Form1.WindowState = 0 Then Dialog.Show
If Form1.Width > 9765 And Form1.WindowState = 1 Then Dialog.Hide
If Form1.Width > 9765 And Form1.WindowState = 0 Then Dialog.Hide
If Form1.Width < 9765 Then Text2.Visible = True
If Form1.Width > 9765 Then Text2.Visible = False
If Form1.WindowState = 2 Then Label45.Visible = False
Label45.Width = Form1.Width - 1695

RichTextBox1.Width = Form1.Width
RichTextBox1.Height = Form1.Height - 2070

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
Label44.Width = Form1.Width - 1695

If RichTextBox1.Top = 1200 Then Frame6.Top = Form1.Height - 645
Frame5.Left = Form1.Width - 3360
Label66.Left = Form1.Width - 360
Label68.Left = Form1.Width - 360
List1.Height = Form1.Height - 2055
Frame7.Left = Form1.Width - 2910
Frame7.Height = Form1.Height - 1935
List2.Height = Frame7.Height - 3360
Frame8.Left = Form1.Width - 2910
Frame8.Height = Form1.Height - 1935
Frame9.Height = Form1.Height - 1935
Frame9.Left = Form1.Width - 3465
List3.Height = Frame9.Height - 2460
Frame6.Top = Form1.Height - 630
Frame10.Top = Form1.Height - 630
Frame11.Top = Form1.Height - 630
Frame12.Top = Form1.Height - 630

If Form1.WindowState = 0 Then List1.Height = Form1.Height - 2070

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload frmOptions
Unload frmSplash
Unload Dialog
End Sub

Private Sub Frame1_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Tahoma"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelColor = RGB(0, 0, 0)
Line6.Visible = True
Line7.Visible = False
Line18.Visible = False
Line19.Visible = False
End Sub

Private Sub Frame10_Click()
Dim i As Long
Dim Counts As Integer
If RichTextBox1.Text = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(RichTextBox1.Text)
        If Mid(RichTextBox1.Text, i, 1) = " " Then
            Counts = Counts + 1
        End If
    Next
End If
Frame10.Caption = Counts
End Sub

Private Sub Frame11_Click()
Dim Counts As Integer
Dim i As Integer
If RichTextBox1.Text = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(RichTextBox1.Text)
        If Mid(RichTextBox1.Text, i, 1) = " " Then Counts = Counts + 1
        
    Next
End If
Frame11.Caption = "Word count: " & Counts
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)
RichTextBox1.SelFontName = "Arial"
RichTextBox1.SelFontSize = "28"
RichTextBox1.SelColor = RGB(86, 163, 233)
RichTextBox1.SelBold = True
Line6.Visible = False
Line7.Visible = True
Line18.Visible = False
Line19.Visible = False
End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Tahoma"
RichTextBox1.SelFontSize = "9"
RichTextBox1.SelColor = &H808080
Line6.Visible = False
Line7.Visible = False
Line18.Visible = True
Line19.Visible = False
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Palatino Linotype"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelColor = RGB(0, 0, 0)
Line6.Visible = False
Line7.Visible = False
Line18.Visible = False
Line19.Visible = True
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
If Line1.Visible = True Then newborderstyle
End Sub

Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line23.Visible = False
Line24.Visible = False
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
Line23.Visible = False
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

Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
End Sub

Private Sub Image31_Click()
 RichTextBox1.SelRTF = RichTextBox2.TextRTF
End Sub

Private Sub Image31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line11.Visible = True
End Sub

Private Sub Image32_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox2.Text = ""
RichTextBox2.TextRTF = RichTextBox1.SelRTF
End Sub

Private Sub Image32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line12.Visible = True
End Sub

Private Sub Image33_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox2.Text = ""
RichTextBox2.TextRTF = RichTextBox1.SelRTF
RichTextBox1.SelRTF = ""
End Sub

Private Sub Image33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line13.Visible = True
End Sub

Private Sub Image34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Line16.Visible = False
Line17.Visible = False
End Sub

Private Sub Image4_Click()
ShowColorDialogFore
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
RichTextBox1.Locked = False
Option1.Value = True
Option2.Value = False
RichTextBox1.Text = ""
Form1.Caption = "Starboy Notepad"
Label2.Caption = "Starboy Notepad"
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

Private Sub Image64_Click()
Form1.Caption = "Starboy Notepad"
Label2.Caption = "Starboy Notepad"



Line1.Visible = False


RichTextBox1.Text = ""
RichTextBox1.Visible = True
End Sub

Private Sub Image65_Click()
RichTextBox1.Visible = True

Line1.Visible = False

OpenFile
End Sub

Private Sub Image66_Click()
RichTextBox1.Visible = True

Line1.Visible = False

salva_bxy
End Sub

Private Sub Image67_Click()


Line1.Visible = False
frmSplash.Show
End Sub

Private Sub Image68_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox2.Text = ""
RichTextBox2.TextRTF = RichTextBox1.SelRTF
RichTextBox1.SelRTF = ""


RichTextBox1.Visible = True



End Sub

Private Sub Image69_Click()

RichTextBox1.Visible = True

Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText, vbCFRTF
RichTextBox2.Text = ""
RichTextBox2.TextRTF = RichTextBox1.SelRTF
End Sub

Private Sub Image7_Click()
FontOpen
End Sub

Private Sub Image70_Click()

 RichTextBox1.SelRTF = RichTextBox2.TextRTF
End Sub

Private Sub Image71_Click()

SendKeys "^Z"
End Sub

Private Sub Image72_Click()

SendKeys "^Y"
End Sub

Private Sub Image73_Click()
AddFile

End Sub

Private Sub Image74_Click()
Dialog.Show

End Sub

Private Sub Image75_Click()

RichTextBox1.SelBullet = 1
Dialog.Image21.Visible = False
Dialog.Image26.Visible = True
End Sub

Private Sub Image76_Click()
frmOptions.Show

End Sub

Private Sub Image77_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image77.Visible = False
End Sub

Private Sub Image78_Click()
Label12.Visible = True
RichTextBox1.Visible = True
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
RichTextBox1.Visible = True
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

Private Sub Image80_Click()
Label12.Visible = True
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
ShowColorDialogFore
End Sub

Private Sub Image82_Click()
ShowColorDialog
End Sub

Private Sub Image83_Click()
Spellcheck
Image77.Visible = True
End Sub

Private Sub Image85_Click()
SendKeys "^A"
frmFind.Show
End Sub

Private Sub Image86_DblClick()
Unload Form1
Unload Dialog
Unload Dialog1
Unload frmFind
Unload frmOptions
Unload frmSplash
End Sub

Private Sub Image87_Click()
VB6IHateYou
End Sub

Private Sub Image87_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line23.Visible = True
Line24.Visible = False
End Sub

Private Sub Image88_Click()
Frame8.Visible = True
End Sub

Private Sub Image88_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line24.Visible = True
Line23.Visible = False
End Sub

Private Sub Image9_Click()
frmOptions.Show
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image29.Visible = True
End Sub


Private Sub Image90_Click()

If Check1.Value = 1 Then List3.AddItem (Text5.Text + " " + "[" + Text6.Text + "/" + Text7.Text + "/" + Text8.Text + "]")
If Check1.Value = 0 Then List3.AddItem (Text5.Text)

Text5.Text = "New task..."
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Image92_Click()
Frame9.Visible = True
End Sub

Private Sub Label13_Click()

Image41.Visible = False
Line1.Visible = True
Image41.Top = 240
Image5.Top = 510
Line8.Y1 = 960
Line8.Y2 = 960
Label52.Top = 1125
Image30.Top = 315
Image13.Top = 450
Line9.Y1 = 690
Line9.Y2 = 690
Image88.Top = 450
Line24.Y1 = 690
Line24.Y2 = 690
Image25.Top = 795
Line10.Y1 = 1005
Line10.Y2 = 1005
Image87.Top = 795
Line23.Y1 = 1005
Line23.Y2 = 1005
Image34.Top = 315
Label51.Top = 1125
Image36.Top = 510
Line16.Y1 = 960
Line16.Y2 = 960
Image33.Top = 450
Line13.Y1 = 690
Line13.Y2 = 690
Image31.Top = 450
Line11.Y1 = 705
Line11.Y2 = 705
Image32.Top = 795
Line12.Y1 = 1005
Line12.Y2 = 1005
Image37.Top = 795
Line17.Y1 = 1005
Line17.Y2 = 1005
Label77.Top = 480
Image38.Top = 315
Label50.Top = 1125
Image7.Top = 450
Image1.Top = 405
Text3.Top = 450
Line14.Y1 = 690
Line14.Y2 = 690
Image6.Top = 795
Image2.Top = 795
Text1.Top = 795
Line15.Y1 = 1020
Line15.Y2 = 1020
Image16.Top = 795
Image15.Top = 795
Image18.Top = 795
Image17.Top = 795
Image20.Top = 795
Image19.Top = 795
Image39.Top = 315
Label49.Top = 1125
Image24.Top = 555
Image10.Top = 555
Image22.Top = 555
Image12.Top = 555
Image23.Top = 555
Image11.Top = 555
Image8.Top = 315
Label48.Top = 1125
Image26.Top = 450
Image21.Top = 450
Label1.Top = 465
Image14.Top = 795
Label9.Top = 795
Image40.Top = 315
Label47.Top = 1125
Label5.Top = 465
Image3.Top = 480
Label6.Top = 810
Image4.Top = 810
Image84.Top = 315
Line3.Y1 = 1320
Line3.Y2 = 360
Line4.Y1 = 1320
Line4.Y2 = 360
Label53.Top = 1125
Frame1.Top = 360
Frame2.Top = 360
Frame3.Top = 360
Frame4.Top = 360
Image27.Top = 480
Label46.Top = 405
Image77.Top = 480
Image83.Top = 480
Label7.Top = 840
Image29.Top = 915
Image9.Top = 915
Image35.Top = 915
Image85.Top = 915
Image92.Top = 915
Label66.Top = 1200
Label68.Top = 1185
Frame5.Top = 1320
Frame7.Top = 1320
Frame8.Top = 1320
Frame9.Top = 1320
RichTextBox1.Top = 1440
List1.Top = 1440
Image41.Top = 240


newborderstyle
End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label15_Click()
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line14.Visible = False
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
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

Private Sub Label44_DblClick()
Form1.WindowState = 0
Label45.Visible = True
End Sub

Private Sub Label45_DblClick()
Form1.WindowState = 2
End Sub

Private Sub Label45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
    
If Form1.WindowState = 0 And Form1.Top <= 0 And Not Form1.Left > 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 And Form1.Top > 0 And Not Form1.Left > 0 Then Form1.WindowState = 0
End Sub

Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line14.Visible = False
 Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX - 1000
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
    
If Form1.WindowState = 0 And Form1.Top <= 0 And Not Form1.Left > 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 And Form1.Top > 0 And Not Form1.Left > 0 Then Form1.WindowState = 0
End Sub

Private Sub Label45_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Form1.WindowState = 0 And Form1.Top <= 0 And Not Form1.Left > 0 Then Form1.WindowState = 2
If Form1.WindowState = 2 And Form1.Top > 0 And Not Form1.Left > 0 Then Form1.WindowState = 0
End Sub

Private Sub Label46_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image77.Visible = True
End Sub

Private Sub Label54_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Tahoma"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelColor = RGB(0, 0, 0)
Line6.Visible = True
Line7.Visible = False
Line18.Visible = False
Line19.Visible = False
End Sub

Private Sub Label55_Click()
RichTextBox1.SelFontName = "Arial"
RichTextBox1.SelFontSize = "28"
RichTextBox1.SelColor = RGB(86, 163, 233)
RichTextBox1.SelBold = True
Line6.Visible = False
Line7.Visible = True
Line18.Visible = False
Line19.Visible = False
End Sub

Private Sub Label56_Click()
RichTextBox1.SelFontName = "Tahoma"
RichTextBox1.SelFontSize = "9"
RichTextBox1.SelColor = &H808080
RichTextBox1.SelBold = False
Line6.Visible = False
Line7.Visible = False
Line18.Visible = True
Line19.Visible = False
End Sub

Private Sub Label57_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Palatino Linotype"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelColor = RGB(0, 0, 0)
Line6.Visible = False
Line7.Visible = False
Line18.Visible = False
Line19.Visible = True
End Sub

Private Sub Label58_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Tahoma"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelColor = RGB(0, 0, 0)
Line6.Visible = True
Line7.Visible = False
Line18.Visible = False
Line19.Visible = False
End Sub

Private Sub Label59_Click()
RichTextBox1.SelFontName = "Arial"
RichTextBox1.SelFontSize = "28"
RichTextBox1.SelColor = RGB(86, 163, 233)
RichTextBox1.SelBold = True
Line6.Visible = False
Line7.Visible = True
Line18.Visible = False
Line19.Visible = False
End Sub

Private Sub Label60_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Tahoma"
RichTextBox1.SelFontSize = "9"
RichTextBox1.SelColor = &H808080
Line6.Visible = False
Line7.Visible = False
Line18.Visible = True
Line19.Visible = False
End Sub

Private Sub Label61_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelFontName = "Palatino Linotype"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelColor = RGB(0, 0, 0)
Line6.Visible = False
Line7.Visible = False
Line18.Visible = False
Line19.Visible = True
End Sub

Private Sub Label66_Click()
Label66.Visible = False
Frame5.Visible = False
Label68.Visible = True
End Sub

Private Sub Label68_Click()
Label66.Visible = True
Frame5.Visible = True
Label68.Visible = False
If Frame11.Visible = True Then Check5.Value = 1
If Frame11.Visible = False Then Check5.Value = 0
If Frame10.Visible = True Then Check6.Value = 1
If Frame10.Visible = False Then Check6.Value = 0
If Frame12.Visible = True Then Check4.Value = 1
If Frame12.Visible = False Then Check4.Value = 0
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = True
Image29.Visible = False
End Sub

Private Sub Label70_Click()
Label66.Visible = False
Label68.Visible = True
Frame5.Visible = False
End Sub

Private Sub Label73_Click()
Dialog1.Show
End Sub

Private Sub Label74_Click()
List2.AddItem (RichTextBox2.Text)
End Sub

Private Sub Label75_Click()
Dim i As Integer
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) Then RichTextBox1.SelText = List2.List(i)
    Next
End Sub

Private Sub Label76_Click()
Frame7.Visible = False
End Sub

Private Sub Label77_Click()
Frame7.Visible = True
End Sub


Private Sub Label82_Click()
Clipboard.SetText (Label81.Caption)
End Sub

Private Sub Label83_Click()
Frame8.Visible = False
End Sub

Private Sub Label85_Click()
Frame9.Visible = False
End Sub

Private Sub Label86_Click()
Dim ind As Integer
ind = List3.ListIndex
If ind >= 0 Then
List3.RemoveItem ind
End If
End Sub

Private Sub Label87_Click()
If MonthView1.Visible = True Then MonthView1.Visible = False
If MonthView1.Visible = False Then MonthView1.Visible = True

End Sub

Private Sub Label9_Click()
AddFile
End Sub

Private Sub Newstyle_Click(Index As Integer)
newborderstyle
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text7.Text = MonthView1.Month
Text8.Text = MonthView1.Year
Text6.Text = MonthView1.Day
MonthView1.Visible = False
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then RichTextBox1.Locked = False
If Option2.Value = True Then RichTextBox1.Locked = True
End Sub

Private Sub Option2_Click()
If Option1.Value = True Then Frame6.Visible = False
If Option2.Value = True Then Frame6.Visible = True
If Option1.Value = True Then RichTextBox1.Locked = False
If Option2.Value = True Then RichTextBox1.Locked = True
End Sub

Private Sub paste_Click()
 RichTextBox1.SelRTF = RichTextBox2.TextRTF
End Sub

Private Sub redo_Click()
SendKeys "^Y"
End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then getemboys

End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuPUForm

    End If

End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
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

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Form2.PopupMenu mnuPUForm
End If
End Sub

Private Sub spell_Click()
Spellcheck
End Sub

Private Sub tasks_Click()
Frame9.Visible = True
End Sub

Private Sub Text1_Change()
On Error Resume Next
RichTextBox1.SelFontSize = Text1.Text

End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
RichTextBox1.SelFontSize = Text1.Text

End Sub

Private Sub Text1_LostFocus()
On Error Resume Next
RichTextBox1.SelFontSize = Text1.Text

End Sub

Private Sub RichTextBox1_Change()

On Error Resume Next
Dim Counts As Integer
Dim i As Integer
If RichTextBox1.Text = "" Then
    Counts = 0
Else
    Counts = 1
    For i = 1 To Len(RichTextBox1.Text)
        If Mid(RichTextBox1.Text, i, 1) = " " Then Counts = Counts + 1
        
    Next
End If
Frame11.Caption = "Word count: " & Counts

Dim Chars As Long

Chars = 0
For i = 0 To Len(RichTextBox1.Text) - 1


    If Asc(RichTextBox1.SelText) > 32 And Asc(RichTextBox1.SelText) <= 126 Then
        Chars = Chars + 1
    End If
Next
Frame10.Caption = "Characters: " & Chars

Frame12.Caption = "Lines:" & Label63.Caption
getemboys
If RichTextBox1.Locked = True Then Option2.Value = True
If RichTextBox1.Locked = False Then Option1.Value = True
List1.FontSize = RichTextBox1.SelFontSize

Dim KeyCode As Integer
Form1.KeyPreview = True
If KeyCode = vbKeyReturn Then getemboys
If KeyCode = vbKeyBack Then getemboys
getemfuckingboys
If RichTextBox1.SelFontName = "Tahoma" And RichTextBox1.SelFontSize = "12" And RichTextBox1.SelColor = RGB(0, 0, 0) Then Line6.Visible = True
If RichTextBox1.SelFontName = "Arial" And RichTextBox1.SelFontSize = "28" And RichTextBox1.SelColor = RGB(86, 163, 233) Then Line7.Visible = True
If RichTextBox1.SelFontName = "Tahoma" And RichTextBox1.SelFontSize = "9" And RichTextBox1.SelColor = &H808080 Then Line18.Visible = True
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


Private Sub Text5_GotFocus()
If Text5.Text = "New task..." Then Text5.Text = ""
If Text5.Text = "New task..." Then Text5.Font.Italic = False
End Sub

Private Sub Text5_LostFocus()
If Text5.Text = "" Then Text5.Text = "New task..."
If Text5.Text = "" Then Text5.Font.Italic = True
End Sub

Private Sub undo_Click()
SendKeys "^Z"
End Sub
