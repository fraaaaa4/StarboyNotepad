VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Options"
   ClientHeight    =   11280
   ClientLeft      =   2580
   ClientTop       =   1515
   ClientWidth     =   12735
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   12735
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Starboy..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   840
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Big thanks to:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         TabIndex        =   64
         Top             =   6240
         Width           =   9015
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Henkkkkk#2077, for helping with Coding, design, ideas and feedback. Creator of Ecstasy Editor, check it out too!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   840
            TabIndex        =   65
            Top             =   360
            Width           =   7815
         End
         Begin VB.Image Image8 
            Height          =   495
            Left            =   240
            Picture         =   "frmOptions.frx":06EA
            Stretch         =   -1  'True
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1320
         TabIndex        =   61
         Top             =   4920
         Width           =   9015
         Begin VB.Image Image14 
            Height          =   375
            Left            =   6360
            Picture         =   "frmOptions.frx":E4075
            Stretch         =   -1  'True
            Top             =   720
            Width           =   1455
         End
         Begin VB.Image Image13 
            Height          =   375
            Left            =   6360
            Picture         =   "frmOptions.frx":10A0BB
            Stretch         =   -1  'True
            Top             =   300
            Width           =   1455
         End
         Begin VB.Image Image12 
            Height          =   855
            Left            =   7920
            Picture         =   "frmOptions.frx":1A7105
            Stretch         =   -1  'True
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Build 1545.srv00_FE.090821-0036/32-public (32-bit First Edition)"
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
            Left            =   360
            TabIndex        =   63
            Top             =   720
            Width           =   8055
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Version 1.53 - 09/08/2021 Update KB75"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   62
            Top             =   480
            Width           =   8295
         End
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmOptions.frx":1CCD67
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   66
         Top             =   7560
         Width           =   9015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   1320
         X2              =   10320
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "32-bit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   9240
         TabIndex        =   60
         Top             =   4080
         Width           =   975
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H8000000A&
         Height          =   495
         Left            =   9000
         Shape           =   4  'Rounded Rectangle
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "A simple, but powerful, Notepad and Wordpad combo app. Compatible with almost everything."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   59
         Top             =   4320
         Width           =   8415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Starboy Notepad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   58
         Top             =   3960
         Width           =   5415
      End
      Begin VB.Image Image7 
         Height          =   3180
         Left            =   2280
         Picture         =   "frmOptions.frx":1CCDFA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6915
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Turn visual elements on or off"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5245
      Left            =   840
      TabIndex        =   7
      Top             =   1440
      Width           =   5655
      Begin VB.CheckBox Check17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Title shadow"
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
         Left            =   2280
         TabIndex        =   38
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox Check16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Background"
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
         Left            =   600
         TabIndex        =   37
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CheckBox Check15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Copyright"
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
         Left            =   4080
         TabIndex        =   36
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox Check14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hide button"
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
         Left            =   2280
         TabIndex        =   35
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tip shape"
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
         Left            =   600
         TabIndex        =   34
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tip line"
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
         Left            =   4080
         TabIndex        =   33
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tip text"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tip title"
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
         Left            =   600
         TabIndex        =   31
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "App version"
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
         Left            =   4080
         TabIndex        =   30
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "App title"
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
         Left            =   2280
         TabIndex        =   29
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "App Icon"
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
         Left            =   600
         TabIndex        =   28
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Close"
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
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "This disables/enables the option to close the Splash Screen."
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Open document"
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
         Left            =   2280
         TabIndex        =   9
         ToolTipText     =   "This disables/enables the option to open a document in the Splash Screen"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New document"
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
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "This disables/enables the option to create a new document on the splash screen"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Image Image25 
         Height          =   495
         Left            =   4200
         Picture         =   "frmOptions.frx":35252C
         Stretch         =   -1  'True
         ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image26 
         Height          =   495
         Left            =   4200
         Picture         =   "frmOptions.frx":3813BE
         Stretch         =   -1  'True
         ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   2880
         TabIndex        =   13
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C000&
         X1              =   360
         X2              =   5280
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Splash Screen"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "You can enable or disable certain elements of the interface, such as buttons, text or pictures."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   4095
      End
      Begin VB.Image Image27 
         Height          =   615
         Left            =   360
         Picture         =   "frmOptions.frx":3AF7A0
         Stretch         =   -1  'True
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Splash screen options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7425
      Left            =   6600
      TabIndex        =   26
      Top             =   1440
      Width           =   5655
      Begin VB.ComboBox Combo15 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   5520
         Width           =   1935
      End
      Begin VB.ComboBox Combo13 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   4920
         Width           =   1935
      End
      Begin VB.ComboBox Combo11 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   3720
         Width           =   1935
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Change"
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
         Left            =   4440
         TabIndex        =   69
         Top             =   6840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Change the background shown in the splash screen. Note that the image should be a 24-bit JPG/bmp image."
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
         Left            =   360
         TabIndex        =   68
         Top             =   6360
         Width           =   5055
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   5280
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Copyright text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tip text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tip title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "App version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Image Image9 
         Height          =   470
         Left            =   360
         Picture         =   "frmOptions.frx":3BD682
         Stretch         =   -1  'True
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change the font of the elements in the Splash Screen, such as the title or the tips window."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   27
         Top             =   600
         Width           =   4095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Titlebar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   840
      TabIndex        =   14
      Top             =   6840
      Width           =   5655
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grey"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   1320
         Width           =   1000
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Top             =   1320
         Width           =   1000
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   1320
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   1680
         Width           =   1000
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Light blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   1680
         Width           =   1000
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Orange"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   1680
         Width           =   1000
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purple"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   1000
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   2040
         Width           =   1000
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   2040
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change the colour of the new titlebar in Starboy Notepad. Note that this option applies only when you enable the new window style."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   24
         Top             =   480
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   240
         Picture         =   "frmOptions.frx":3CBD14
         Stretch         =   -1  'True
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3135
      Left            =   11760
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   67
      Top             =   480
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3135
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4575
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   13335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   530
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   600
      Picture         =   "frmOptions.frx":3D8F72
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image15 
      Height          =   5415
      Left            =   0
      Picture         =   "frmOptions.frx":3D965C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   11160
      Picture         =   "frmOptions.frx":82B27B
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   11160
      Picture         =   "frmOptions.frx":84349D
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10920
      TabIndex        =   6
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About..."
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
      Left            =   6600
      TabIndex        =   40
      Top             =   10560
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personalisation"
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
      Left            =   5205
      TabIndex        =   39
      Top             =   10560
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   6840
      Picture         =   "frmOptions.frx":85B6BF
      Stretch         =   -1  'True
      Top             =   10080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   6840
      Picture         =   "frmOptions.frx":868A51
      Stretch         =   -1  'True
      Top             =   10080
      Width           =   435
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   5640
      Picture         =   "frmOptions.frx":875753
      Stretch         =   -1  'True
      Top             =   10080
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   5640
      Picture         =   "frmOptions.frx":8831ED
      Stretch         =   -1  'True
      Top             =   10080
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   9720
      Width           =   12615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   9720
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   10695
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   12375
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Option Explicit

Private Sub cmdApply_Click()
    MsgBox "Place code here to set options w/o closing dialog!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label3.BorderStyle = 1
frmSplash.Label2.BorderStyle = 0
frmSplash.Label1.BorderStyle = 0
End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then frmSplash.Label11.Visible = True
If Check10.Value = 0 Then frmSplash.Label11.Visible = False
End Sub

Private Sub Check11_Click()
If Check11.Value = 1 Then frmSplash.Label12.Visible = True
If Check11.Value = 0 Then frmSplash.Label12.Visible = False
If Check11.Value = 1 Then frmSplash.Label15.Visible = True
If Check11.Value = 0 Then frmSplash.Label15.Visible = False
If Check11.Value = 1 Then frmSplash.Label16.Visible = True
If Check11.Value = 0 Then frmSplash.Label16.Visible = False
End Sub

Private Sub Check12_Click()
If Check12.Value = 1 Then frmSplash.Line1.Visible = True
If Check12.Value = 0 Then frmSplash.Line1.Visible = False
End Sub

Private Sub Check13_Click()
If Check13.Value = 1 Then frmSplash.Shape7.Visible = True
If Check13.Value = 0 Then frmSplash.Shape7.Visible = False
End Sub

Private Sub Check14_Click()
If Check14.Value = 1 Then frmSplash.Label14.Visible = True
If Check14.Value = 0 Then frmSplash.Label14.Visible = False
End Sub

Private Sub Check15_Click()
If Check15.Value = 1 Then frmSplash.Label10.Visible = True
If Check15.Value = 0 Then frmSplash.Label10.Visible = False
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then frmSplash.Image1.Visible = True
If Check16.Value = 0 Then frmSplash.Image1.Visible = False
End Sub

Private Sub Check17_Click()
If Check17.Value = 1 Then frmSplash.Label4.Visible = True
If Check17.Value = 0 Then frmSplash.Label4.Visible = False
End Sub

Private Sub Check2_Click()

End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label2.BorderStyle = 1
frmSplash.Label3.BorderStyle = 0
frmSplash.Label1.BorderStyle = 0
End Sub

Private Sub Check3_Click()

End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label1.BorderStyle = 1
frmSplash.Label3.BorderStyle = 0
frmSplash.Label2.BorderStyle = 0

End Sub

Private Sub Command1_Click()

End Sub


Private Sub Check4_Click()
If Check4.Value = 1 Then frmSplash.Label3.Visible = True
If Check4.Value = 0 Then frmSplash.Label3.Visible = False
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then frmSplash.Label2.Visible = True
If Check5.Value = 0 Then frmSplash.Label2.Visible = False
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then frmSplash.Label1.Visible = True
If Check6.Value = 0 Then frmSplash.Label1.Visible = False
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then frmSplash.Image2.Visible = True
If Check7.Value = 0 Then frmSplash.Image2.Visible = False
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then frmSplash.Label5.Visible = True
If Check8.Value = 0 Then frmSplash.Label5.Visible = False
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then frmSplash.Label6.Visible = True
If Check9.Value = 0 Then frmSplash.Label6.Visible = False
End Sub


Private Sub Combo1_Change()
frmSplash.Label5.Font = Combo1.Text
frmSplash.Label4.Font = Combo1.Text
End Sub

Private Sub Combo1_Click()
frmSplash.Label5.Font = Combo1.Text
frmSplash.Label4.Font = Combo1.Text
End Sub

Private Sub Combo11_Change()
frmSplash.Label11.Font = Combo11.Text
End Sub

Private Sub Combo11_Click()
frmSplash.Label11.Font = Combo11.Text
End Sub

Private Sub Combo13_Change()
frmSplash.Label12.Font = Combo13.Text
frmSplash.Label15.Font = Combo13.Text
frmSplash.Label16.Font = Combo13.Text
End Sub

Private Sub Combo13_Click()
frmSplash.Label12.Font = Combo13.Text
frmSplash.Label15.Font = Combo13.Text
frmSplash.Label16.Font = Combo13.Text
End Sub

Private Sub Combo15_Change()
frmSplash.Label10.Font = Combo15.Text
End Sub

Private Sub Combo15_Click()
frmSplash.Label10.Font = Combo15.Text
End Sub

Private Sub Combo3_Change()
frmSplash.Label6.Font = Combo3.Text
End Sub

Private Sub Combo3_Click()
frmSplash.Label6.Font = Combo3.Text
End Sub

Private Sub Combo5_Change()
frmSplash.Label7.Font = Combo5.Text
End Sub

Private Sub Combo5_Click()
frmSplash.Label7.Font = Combo5.Text
End Sub

Private Sub Combo7_Change()
frmSplash.Label8.Font = Combo7.Text
End Sub

Private Sub Combo7_Click()
frmSplash.Label8.Font = Combo7.Text
End Sub

Private Sub Combo9_Change()
frmSplash.Label9.Font = Combo9.Text
End Sub

Private Sub Combo9_Click()
frmSplash.Label9.Font = Combo9.Text
End Sub

Private Sub Form_Load()
For i = 0 To Screen.FontCount - 1
Combo1.AddItem Screen.Fonts(i)
Combo3.AddItem Screen.Fonts(i)
Combo5.AddItem Screen.Fonts(i)
Combo7.AddItem Screen.Fonts(i)
Combo9.AddItem Screen.Fonts(i)
Combo11.AddItem Screen.Fonts(i)
Combo13.AddItem Screen.Fonts(i)
Combo15.AddItem Screen.Fonts(i)
Next

Combo1.Text = frmSplash.Label5.Font
Combo3.Text = frmSplash.Label6.Font
Combo5.Text = frmSplash.Label7.Font
Combo7.Text = frmSplash.Label8.Font
Combo9.Text = frmSplash.Label9.Font
Combo11.Text = frmSplash.Label11.Font
Combo13.Text = frmSplash.Label12.Font
Combo13.Text = frmSplash.Label15.Font
Combo13.Text = frmSplash.Label16.Font
Combo15.Text = frmSplash.Label10.Font

Image6.Visible = False
Image4.Visible = False
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Form1.Show
If frmSplash.Label2.Visible = True Then Check5.Value = 1
If frmSplash.Label2.Visible = False Then Check5.Value = 0
If frmSplash.Label1.Visible = True Then Check6.Value = 1
If frmSplash.Label1.Visible = False Then Check6.Value = 0
If frmSplash.Label1.Visible = True Then Check4.Value = 1
If frmSplash.Label1.Visible = False Then Check4.Value = 0
If frmSplash.Label11.Visible = True Then Check10.Value = 1
If frmSplash.Label11.Visible = False Then Check10.Value = 0
If frmSplash.Image2.Visible = True Then Check7.Value = 1
If frmSplash.Image2.Visible = False Then Check7.Value = 0
If frmSplash.Label5.Visible = True Then Check8.Value = 1
If frmSplash.Label5.Visible = False Then Check8.Value = 0
If frmSplash.Label4.Visible = True Then Check17.Value = 1
If frmSplash.Label4.Visible = False Then Check17.Value = 0
If frmSplash.Label6.Visible = True Then Check9.Value = 1
If frmSplash.Label6.Visible = False Then Check9.Value = 0
If frmSplash.Label11.Visible = True Then Check10.Value = 1
If frmSplash.Label11.Visible = False Then Check10.Value = 0
If frmSplash.Line1.Visible = True Then Check12.Value = 1
If frmSplash.Line1.Visible = False Then Check12.Value = 0
If frmSplash.Label2.Visible = True Then Check11.Value = 1
If frmSplash.Label12.Visible = True Then Check11.Value = 0
If frmSplash.Shape7.Visible = True Then Check13.Value = 1
If frmSplash.Shape7.Visible = False Then Check13.Value = 0
If frmSplash.Label14.Visible = True Then Check14.Value = 1
If frmSplash.Label14.Visible = False Then Check14.Value = 0
If frmSplash.Label10.Visible = True Then Check15.Value = 1
If frmSplash.Label10.Visible = False Then Check15.Value = 0
If frmSplash.Image1.Visible = True Then Check16.Value = 1
If frmSplash.Image1.Visible = False Then Check16.Value = 0
If Form1.Image42.Visible = True Then Option1.Value = True
If Form1.Image42.Visible = False Then Option1.Value = False
If Form1.Image49.Visible = True Then Option2.Value = True
If Form1.Image49.Visible = False Then Option2.Value = False
If Form1.Image52.Visible = True Then Option3.Value = True
If Form1.Image52.Visible = False Then Option3.Value = False
If Form1.Image57.Visible = True Then Option4.Value = True
If Form1.Image57.Visible = False Then Option4.Value = False
If Form1.Image58.Visible = True Then Option5.Value = True
If Form1.Image58.Visible = False Then Option5.Value = False
If Form1.Image59.Visible = True Then Option6.Value = True
If Form1.Image59.Visible = False Then Option6.Value = False
If Form1.Image60.Visible = True Then Option7.Value = True
If Form1.Image60.Visible = False Then Option7.Value = False
If Form1.Image61.Visible = True Then Option8.Value = True
If Form1.Image61.Visible = False Then Option8.Value = False
If Form1.Image62.Visible = True Then Option9.Value = True
If Form1.Image62.Visible = False Then Option9.Value = False





End Sub

Private Sub Form_Resize()
Image11.Top = frmOptions.Height - 1485
Image10.Top = frmOptions.Height - 1485
Label7.Top = frmOptions.Height - 1605
Label10.Top = frmOptions.Height - 1125
Label9.Top = frmOptions.Height - 1125
Image4.Top = frmOptions.Height - 1635
Image6.Top = frmOptions.Height - 1635
Image3.Top = frmOptions.Height - 1635
Image5.Top = frmOptions.Height - 1635
Shape3.Top = frmOptions.Height - 1995
Shape4.Top = frmOptions.Height - 1875
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
frmOptions.Hide
frmSplash.Hide
End Sub

Private Sub fraSample1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label3.BorderStyle = 0
frmSplash.Label2.BorderStyle = 0
frmSplash.Label1.BorderStyle = 0
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
If Frame1.Height = 6225 Then Frame1.Height = 345
If Frame1.Height = 345 Then Frame1.Height = 6225

End Sub

Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)
If Frame6.Height = 2625 Then Frame6.Height = 345
If Frame6.Height = 345 Then Frame6.Height = 2625
End Sub

Private Sub Frame7_Click()
If Frame7.Height = 5145 Then Frame7.Height = 345
If Frame7.Height = 345 Then Frame7.Height = 5145
If Frame7.Height = 5145 Then Frame6.Top = 6840
If Frame7.Height = 345 Then Frame6.Top = 1920
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmOptions.Hide
Form1.Show
End Sub

Private Sub Image11_Click()
frmOptions.Hide
Form1.Show
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmOptions.Hide
Form1.Show


End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = False
Image10.Visible = True

End Sub

Private Sub Image16_Click()
fuck
End Sub

Private Sub Image19_Click()
Label1.Font = Form1.CommonDialog1.FontName
Label1.Font.Size = Form1.CommonDialog1.FontSize
Label1.Font.Bold = Form1.CommonDialog1.FontBold
Label1.Font.Italic = Form1.CommonDialog1.FontItalic
Label1.Font.Underline = Form1.CommonDialog1.FontUnderline
End Sub

Private Sub Image22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Image24_Click()

End Sub

Private Sub Image25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image26.Visible = True
Image25.Visible = False
End Sub

Private Sub Image26_Click()
frmSplash.Show
frmSplash.Left = frmOptions.Left + 100
frmSplash.Top = frmOptions.Top + 100
End Sub

Private Sub Image3_Click()
Shape6.Visible = True
Shape1.Visible = True
Frame7.Visible = True
Frame1.Visible = True
Frame6.Visible = True
Label10.Font.Bold = False
Label9.Font.Bold = True
Image3.Visible = False
Image5.Visible = True
Image6.Visible = False
Image4.Visible = True
Frame2.Visible = False
End Sub

Private Sub Image5_Click()
Shape6.Visible = False
Shape1.Visible = False
Frame7.Visible = False
Frame1.Visible = False
Frame6.Visible = False
Label9.Font.Bold = False
Label10.Font.Bold = True
Image5.Visible = False
Image3.Visible = True
Image4.Visible = False
Image6.Visible = True
Frame2.Visible = True
End Sub

Private Sub Image8_Click()
frmOptions.Hide
Form1.Show
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.Visible = True
Image26.Visible = False
End Sub

Private Sub Label29_Click()
End Sub

Private Sub Label26_Click()
If Frame7.Height = 5145 Then Frame7.Height = 345
If Frame7.Height = 345 Then Frame7.Height = 5145
If Frame7.Height = 5145 Then Frame6.Top = 6840
If Frame7.Height = 345 Then Frame6.Top = 1920
End Sub

Private Sub Label4_Click()
CommonDialog1.InitDir = App.Path
        CommonDialog1.filename = ""
        CommonDialog1.Filter = "JPEG Image (*.jpg)|*.jpg|BMP Image (*.bmp)|*,bmp|All Files (*.*)|*.*"
        CommonDialog1.DialogTitle = "Open Image"
        CommonDialog1.ShowOpen
 
        If CommonDialog1.filename <> "" Then
            frmSplash.Image1.Picture = LoadPicture(CommonDialog1.filename)
End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image4.Visible = False
End Sub

Private Sub Label7_Click()
Image11.Visible = True
Image10.Visible = False
1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image10.Visible = False

End Sub

Private Sub Option1_Click()
Form1.Image43.Visible = True
Form1.Image42.Visible = True
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
End Sub

Private Sub Option10_Click()
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
Form1.Image56.Visible = True
Form1.Image62.Visible = True
End Sub

Private Sub Option11_Click()
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
Form1.Image55.Visible = True
Form1.Image61.Visible = True
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option12_Click()
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
Form1.Image54.Visible = True
Form1.Image60.Visible = True
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option13_Click()
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
Form1.Image53.Visible = True
Form1.Image59.Visible = True
Form1.Image54.Visible = False
Form1.Image60.Visible = False
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option14_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = False
Form1.Image52.Visible = False
Form1.Image47.Visible = False
Form1.Image57.Visible = False
Form1.Image50.Visible = True
Form1.Image58.Visible = True
Form1.Image53.Visible = False
Form1.Image59.Visible = False
Form1.Image54.Visible = False
Form1.Image60.Visible = False
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option15_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = False
Form1.Image52.Visible = False
Form1.Image47.Visible = True
Form1.Image57.Visible = True
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
End Sub

Private Sub Option16_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = True
Form1.Image52.Visible = True
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
End Sub

Private Sub Option17_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = True
Form1.Image49.Visible = True
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
End Sub

Private Sub Option18_Click()
Form1.Image43.Visible = True
Form1.Image42.Visible = True
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
End Sub

Private Sub Option2_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = True
Form1.Image49.Visible = True
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
End Sub

Private Sub Option3_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = True
Form1.Image52.Visible = True
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
End Sub

Private Sub Option4_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = False
Form1.Image52.Visible = False
Form1.Image47.Visible = True
Form1.Image57.Visible = True
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
End Sub

Private Sub Option5_Click()
Form1.Image43.Visible = False
Form1.Image42.Visible = False
Form1.Image48.Visible = False
Form1.Image49.Visible = False
Form1.Image51.Visible = False
Form1.Image52.Visible = False
Form1.Image47.Visible = False
Form1.Image57.Visible = False
Form1.Image50.Visible = True
Form1.Image58.Visible = True
Form1.Image53.Visible = False
Form1.Image59.Visible = False
Form1.Image54.Visible = False
Form1.Image60.Visible = False
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option6_Click()
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
Form1.Image53.Visible = True
Form1.Image59.Visible = True
Form1.Image54.Visible = False
Form1.Image60.Visible = False
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option7_Click()
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
Form1.Image54.Visible = True
Form1.Image60.Visible = True
Form1.Image55.Visible = False
Form1.Image61.Visible = False
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option8_Click()
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
Form1.Image55.Visible = True
Form1.Image61.Visible = True
Form1.Image56.Visible = False
Form1.Image62.Visible = False
End Sub

Private Sub Option9_Click()
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
Form1.Image56.Visible = True
Form1.Image62.Visible = True
End Sub

Private Sub RichTextBox2_Change()

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Text2_Change()

End Sub
