VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7860
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5685
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "This menu changes how the titlebar looks"
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Visual effects"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "fraSample1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "About...."
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Image12"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Image13"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Image14"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame3"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Fonts"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Label16"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Titlebar"
      TabPicture(3)   =   "frmOptions.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label18"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Splash Screen Font"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   -75000
         TabIndex        =   48
         Top             =   3360
         Width           =   5655
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   615
            Left            =   480
            TabIndex        =   49
            Top             =   1440
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1085
            _Version        =   393217
            BorderStyle     =   0
            Appearance      =   0
            TextRTF         =   $"frmOptions.frx":007C
         End
         Begin VB.Image Image23 
            Height          =   615
            Left            =   360
            Picture         =   "frmOptions.frx":0114
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label21 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(To be moved in the new settings) Change the font used in the Splash Screen Title. Default is Tahoma "
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
            TabIndex        =   50
            Top             =   600
            Width           =   4095
         End
         Begin VB.Image Image22 
            Height          =   495
            Left            =   4320
            Picture         =   "frmOptions.frx":1F106
            Stretch         =   -1  'True
            ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
            Top             =   1320
            Width           =   855
         End
         Begin VB.Image Image20 
            Height          =   855
            Left            =   360
            Picture         =   "frmOptions.frx":4EE30
            Stretch         =   -1  'True
            Top             =   1275
            Width           =   3735
         End
         Begin VB.Image Image24 
            Height          =   495
            Left            =   4320
            Picture         =   "frmOptions.frx":B4072
            Stretch         =   -1  'True
            ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   3000
            TabIndex        =   51
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
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
         Left            =   -75000
         TabIndex        =   37
         Top             =   600
         Width           =   5655
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   2040
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
            TabIndex        =   45
            Top             =   2040
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   1680
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
            TabIndex        =   42
            Top             =   1680
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
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   1320
            Width           =   1000
         End
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
            TabIndex        =   39
            Top             =   1320
            Width           =   1000
         End
         Begin VB.Image Image19 
            Height          =   495
            Left            =   240
            Picture         =   "frmOptions.frx":E2BF4
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label20 
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
            TabIndex        =   38
            Top             =   480
            Width           =   4455
         End
      End
      Begin VB.Frame fraSample1 
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
         Height          =   3345
         Left            =   -75000
         TabIndex        =   9
         Top             =   600
         Width           =   5655
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   600
            TabIndex        =   12
            ToolTipText     =   "This disables/enables the option to create a new document on the splash screen"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            ToolTipText     =   "This disables/enables the option to open a document in the Splash Screen"
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Left            =   4080
            TabIndex        =   10
            ToolTipText     =   "This disables/enables the option to close the Splash Screen."
            Top             =   1800
            Width           =   855
         End
         Begin VB.Image Image3 
            Height          =   495
            Left            =   4200
            Picture         =   "frmOptions.frx":EFE52
            Stretch         =   -1  'True
            ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
            Top             =   2280
            Width           =   855
         End
         Begin VB.Image Image4 
            Height          =   495
            Left            =   4200
            Picture         =   "frmOptions.frx":11ECE4
            Stretch         =   -1  'True
            ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
            Top             =   2280
            Width           =   855
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   360
            Picture         =   "frmOptions.frx":14D0C6
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
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
            TabIndex        =   15
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label2 
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
            TabIndex        =   14
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C000&
            X1              =   360
            X2              =   5280
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Height          =   1215
            Left            =   3000
            TabIndex        =   13
            Top             =   1800
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Font options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   -75000
         TabIndex        =   16
         Top             =   4200
         Width           =   5655
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
            Left            =   2160
            TabIndex        =   17
            Text            =   "Text1"
            ToolTipText     =   "This changes the font used on the labels on the notepad"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Image Image7 
            Height          =   495
            Left            =   4200
            Picture         =   "frmOptions.frx":15AFA8
            Stretch         =   -1  'True
            ToolTipText     =   "This reverts the option to the default option (Tahoma)"
            Top             =   1200
            Width           =   855
         End
         Begin VB.Image Image6 
            Height          =   495
            Left            =   4200
            Picture         =   "frmOptions.frx":18ACD2
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Notepad labels"
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
            TabIndex        =   20
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Change the font used by the application for all the text included."
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
            Left            =   1200
            TabIndex        =   19
            Top             =   480
            Width           =   4095
         End
         Begin VB.Image Image2 
            Height          =   375
            Left            =   360
            Picture         =   "frmOptions.frx":1B9854
            Stretch         =   -1  'True
            Top             =   480
            Width           =   560
         End
         Begin VB.Image Image5 
            Height          =   495
            Left            =   2040
            Picture         =   "frmOptions.frx":1C7EE6
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C000&
            X1              =   2160
            X2              =   3600
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   4080
            TabIndex        =   18
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Thanks to:"
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
         Left            =   480
         TabIndex        =   26
         Top             =   3300
         Width           =   4575
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "We would like to specifically Thank You to Hanghitorgame, for helping with the code and other things for this notepad."
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
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   975
         Left            =   480
         TabIndex        =   24
         Top             =   2220
         Width           =   4575
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Starboy Notepad - version 1.3.1.1110. Update version KB41 (17/07/2021)"
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
            TabIndex        =   25
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default font"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   -75000
         TabIndex        =   32
         Top             =   600
         Width           =   5655
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   615
            Left            =   480
            TabIndex        =   35
            Top             =   1440
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1085
            _Version        =   393217
            BorderStyle     =   0
            Appearance      =   0
            TextRTF         =   $"frmOptions.frx":22D128
         End
         Begin VB.Image Image18 
            Height          =   855
            Left            =   360
            Picture         =   "frmOptions.frx":22D1C0
            Stretch         =   -1  'True
            Top             =   1275
            Width           =   3735
         End
         Begin VB.Image Image15 
            Height          =   495
            Left            =   4320
            Picture         =   "frmOptions.frx":292402
            Stretch         =   -1  'True
            ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Choose what is the font that the notepad should use every time you create a new document."
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
            TabIndex        =   33
            Top             =   600
            Width           =   4095
         End
         Begin VB.Image Image17 
            Height          =   615
            Left            =   360
            Picture         =   "frmOptions.frx":2C0344
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image Image16 
            Height          =   495
            Left            =   4320
            Picture         =   "frmOptions.frx":2DF336
            Stretch         =   -1  'True
            ToolTipText     =   "You can preview what you activated or not directly on the splash screen"
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   3000
            TabIndex        =   34
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Height          =   6780
         Left            =   -75000
         TabIndex        =   36
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Height          =   6780
         Left            =   -75000
         TabIndex        =   31
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6795
         Left            =   -75000
         TabIndex        =   30
         Top             =   360
         Width           =   5655
      End
      Begin VB.Image Image14 
         Height          =   495
         Left            =   2400
         Picture         =   "frmOptions.frx":30C948
         Stretch         =   -1  'True
         Top             =   6300
         Width           =   1935
      End
      Begin VB.Image Image13 
         Height          =   615
         Left            =   2400
         Picture         =   "frmOptions.frx":33298E
         Stretch         =   -1  'True
         Top             =   5460
         Width           =   2535
      End
      Begin VB.Image Image12 
         Height          =   1455
         Left            =   600
         Picture         =   "frmOptions.frx":3CF9D8
         Stretch         =   -1  'True
         Top             =   5340
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Completely made in:"
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
         TabIndex        =   29
         Top             =   4980
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright (c) 2014-2021 MoonLight Corp. All rights reserved. Not officially copyrighted"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   4620
         Width           =   4575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version 1.31 - build 1110.public/-a 21717"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   23
         Top             =   1380
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Starboy Notepad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   22
         Top             =   900
         Width           =   2775
      End
      Begin VB.Image Image9 
         Height          =   1095
         Left            =   480
         Picture         =   "frmOptions.frx":3F563A
         Stretch         =   -1  'True
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Height          =   6780
         Left            =   0
         TabIndex        =   21
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   6
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   5
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   0
      Left            =   0
      ScaleHeight     =   5100
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   5685
   End
   Begin VB.Image Image21 
      Height          =   615
      Left            =   0
      Picture         =   "frmOptions.frx":3F5BC4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   4560
      Picture         =   "frmOptions.frx":414BB6
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   4560
      Picture         =   "frmOptions.frx":42CDD8
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   4560
      Picture         =   "frmOptions.frx":444FFA
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3960
      TabIndex        =   7
      Top             =   7080
      Width           =   1575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
If Check1.Value = 1 Then frmSplash.Label3.Visible = True
If Check1.Value = 0 Then frmSplash.Label3.Visible = False
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label3.BorderStyle = 1
frmSplash.Label2.BorderStyle = 0
frmSplash.Label1.BorderStyle = 0
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then frmSplash.Label2.Visible = True
If Check2.Value = 0 Then frmSplash.Label2.Visible = False
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label2.BorderStyle = 1
frmSplash.Label3.BorderStyle = 0
frmSplash.Label1.BorderStyle = 0
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then frmSplash.Label1.Visible = True
If Check3.Value = 0 Then frmSplash.Label1.Visible = False
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label1.BorderStyle = 1
frmSplash.Label3.BorderStyle = 0
frmSplash.Label2.BorderStyle = 0

End Sub

Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Form1.Show
If frmSplash.Label2.Visible = True Then Check2.Value = 1
If frmSplash.Label2.Visible = False Then Check2.Value = 0
If frmSplash.Label1.Visible = True Then Check3.Value = 1
If frmSplash.Label1.Visible = False Then Check3.Value = 0
If frmSplash.Label1.Visible = True Then Check1.Value = 1
If frmSplash.Label1.Visible = False Then Check1.Value = 0
Text1.Text = Form1.Label2.Font
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
RichTextBox2.Text = frmSplash.Label5.Font
RichTextBox2.Text = frmSplash.Label4.Font
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image10.Visible = False
Image8.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
frmOptions.Hide
frmSplash.Hide
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line2.Visible = False
Image11.Visible = True
Image10.Visible = False
Image8.Visible = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.Visible = True
Image16.Visible = False
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = True
Image24.Visible = False
End Sub

Private Sub fraSample1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmSplash.Label3.BorderStyle = 0
frmSplash.Label2.BorderStyle = 0
frmSplash.Label1.BorderStyle = 0
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
Image8.Visible = False
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.Visible = False
Image16.Visible = True
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
Image22.Visible = False
Image24.Visible = True
End Sub

Private Sub Image24_Click()
frmSplash.Label5.Font = "Tahoma"
frmSplash.Label4.Font = "Tahoma"
RichTextBox1.Text = "Tahoma"
End Sub

Private Sub Image3_Click()
frmSplash.Show
frmSplash.Left = 100
frmSplash.Top = 100
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image4.Visible = True
End Sub

Private Sub Image4_Click()
frmSplash.Show
frmSplash.Left = 100
frmSplash.Top = 100
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line2.Visible = False
End Sub


Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image7.Visible = False
End Sub

Private Sub Image8_Click()
frmOptions.Hide
Form1.Show
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image16.Visible = False
Image15.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line2.Visible = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image4.Visible = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image6.Visible = False
End Sub

Private Sub Label7_Click()
Image11.Visible = True
Image10.Visible = False
Image8.Visible = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image10.Visible = False
Image8.Visible = False
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

Private Sub picOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Line2.Visible = False
Image11.Visible = True
Image10.Visible = False
Image8.Visible = False
End Sub


Private Sub RichTextBox2_Change()
frmSplash.Label5.Font = RichTextBox2.Text
frmSplash.Label4.Font = RichTextBox2.Text
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line2.Visible = True
End Sub

Private Sub Text2_Change()

End Sub
