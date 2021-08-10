VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Starboy Installer (32-bit)"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "FormExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16200
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   6960
      Left            =   0
      TabIndex        =   27
      Top             =   -120
      Visible         =   0   'False
      Width           =   7920
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Create a shortcut on desktop"
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
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   6360
         Width           =   2535
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Finish"
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
         Left            =   4800
         TabIndex        =   30
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<          &Back"
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
         Left            =   3240
         TabIndex        =   29
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
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
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   28
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Image Image12 
         Height          =   2535
         Left            =   840
         Picture         =   "FormExample.frx":06EA
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   6135
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "If you encounter any program, or want to submit a suggestion, be sure to check the Issues section on our Github page!"
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
         Height          =   615
         Left            =   360
         TabIndex        =   35
         Top             =   2760
         Width           =   7215
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"FormExample.frx":B6784
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
         Height          =   975
         Left            =   360
         TabIndex        =   33
         Top             =   1920
         Width           =   7095
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   6720
         Picture         =   "FormExample.frx":B6844
         Stretch         =   -1  'True
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "All ready!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "The installation process has finished."
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
         Left            =   360
         TabIndex        =   31
         Top             =   960
         Width           =   5895
      End
      Begin VB.Image Image11 
         Height          =   1335
         Left            =   0
         Picture         =   "FormExample.frx":B6F2E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   6960
      Left            =   0
      TabIndex        =   18
      Top             =   -120
      Visible         =   0   'False
      Width           =   7920
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "I'm running this on Windows Vista or superior"
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
         Left            =   360
         TabIndex        =   26
         Top             =   4080
         Width           =   3735
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next         >"
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
         Left            =   4800
         TabIndex        =   21
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<          &Back"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
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
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   19
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Image Image9 
         Height          =   1575
         Left            =   1800
         Picture         =   "FormExample.frx":2688A0
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   4215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"FormExample.frx":28ED5A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   360
         TabIndex        =   25
         Top             =   2880
         Width           =   7095
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"FormExample.frx":28EEC8
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
         Height          =   735
         Left            =   360
         TabIndex        =   24
         Top             =   1920
         Width           =   7095
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   6720
         Picture         =   "FormExample.frx":28EF93
         Stretch         =   -1  'True
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Installation notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Be sure to read these before starting installation"
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
         Left            =   360
         TabIndex        =   22
         Top             =   960
         Width           =   5895
      End
      Begin VB.Image Image6 
         Height          =   1335
         Left            =   0
         Picture         =   "FormExample.frx":28F67D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6960
      Left            =   0
      TabIndex        =   10
      Top             =   -120
      Visible         =   0   'False
      Width           =   7920
      Begin VB.TextBox Text2 
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
         Height          =   3375
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "FormExample.frx":440FEF
         Top             =   2640
         Width           =   7215
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
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
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   16
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<          &Back"
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
         Left            =   3240
         TabIndex        =   12
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next         >"
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
         Left            =   4800
         TabIndex        =   11
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Starboy Notepad version: KB75 1.53.1545"
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
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Release notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   5175
      End
      Begin VB.Image Image7 
         Height          =   720
         Left            =   6720
         Picture         =   "FormExample.frx":441D43
         Stretch         =   -1  'True
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Here are all the new things we introduced in this version of Starboy Notepad. Be sure to check our Github page for new updates!"
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
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   1920
         Width           =   7095
      End
      Begin VB.Image Image8 
         Height          =   1335
         Left            =   0
         Picture         =   "FormExample.frx":44242D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
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
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6960
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7920
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2655
         Left            =   480
         TabIndex        =   9
         Top             =   3360
         Width           =   7095
         ExtentX         =   12515
         ExtentY         =   4683
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
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
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "%USERPROFILE%"
         Top             =   2760
         Width           =   5295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse..."
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
         Left            =   6360
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next         >"
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
         Left            =   4800
         TabIndex        =   2
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<          &Back"
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
         Left            =   3240
         TabIndex        =   1
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   480
         Picture         =   "FormExample.frx":5F3D9F
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   6855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Select a directory to which install Starboy Notepad. All the files will occupy around 10-20MB of space on your local disk."
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
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   1920
         Width           =   6255
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   480
         Picture         =   "FormExample.frx":658FE1
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   6720
         Picture         =   "FormExample.frx":6596CB
         Stretch         =   -1  'True
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select your installation directory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Starboy Notepad will be installed in the directory chosen by you."
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
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   7215
      End
      Begin VB.Image Image2 
         Height          =   1335
         Left            =   0
         Picture         =   "FormExample.frx":659DB5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload FormWelcome
Unload Form1
End Sub

Private Sub Command10_Click()
Frame3.Visible = False
Frame4.Visible = True
End Sub

Private Sub Command11_Click()
Unload FormWelcome
Unload Form1
End Sub

Private Sub Command12_Click()
Frame4.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command13_Click()
On Error Resume Next
Form2.Show
Form2.Label3.Caption = "Phase 1 of 13: Copying Starboy..."
Form2.ProgressBar1.Value = 1
FileCopy App.Path & "\StarboyKB75.exe", Text1.Text & "\StarboyKB75.exe"
Form2.Label3.Caption = "Phase 2 of 13: Copying Starboy to your Start Menu..."
Form2.ProgressBar1.Value = 2

Set WshShell = WScript.CreateObject("WScript.Shell")
strStart = WshShell.SpecialFolders("StartMenu")
Set oShellLink = WshShell.CreateShortcut(strStart & "\Starboy Notepad.lnk")
oShellLink.TargetPath = Form1.Text1.Text & "\StarboyKB75.exe" 'WScript.ScriptFullName
oShellLink.WindowStyle = 1
oShellLink.Hotkey = "CTRL+SHIFT+N"
oShellLink.IconLocation = Form1.Text1.Text & "\StarboyKB75.exe, 0"
oShellLink.Description = "Shortcut To Starboy Notepad."
oShellLink.WorkingDirectory = strStart
oShellLink.Save

FileCopy App.Path & "\StarboyKB75.exe", "%USERPROFILE%\Start Menu\StarboyKB75.exe"
Form2.Label3.Caption = "Phase 3 of 13: Copying Starboy to your Start Menu..."
Form2.ProgressBar1.Value = 3
FileCopy App.Path & "\StarboyKB75.exe", "%ProgramData%\Microsoft\Windows\Start Menu\Programs\StarboyKB75.exe"
Form2.Label3.Caption = "Phase 4 of 13: Copying starboyreg.bat..."
Form2.ProgressBar1.Value = 4
FileCopy App.Path & "\Starboyreg.bat", "C:\Starboyreg.bat"
Form2.Label3.Caption = "Phase 5 of 13: Copying the first ocx control..."
Form2.ProgressBar1.Value = 5
FileCopy App.Path & "\MSCOMCT2.OCX", "C:\MSCOMCT2.OCX"
Form2.Label3.Caption = "Phase 6 of 13: Copying the second ocx control..."
Form2.ProgressBar1.Value = 6
FileCopy App.Path & "\TABCTL32.OCX", "C:\TABCTL32.OCX"
Form2.Label3.Caption = "Phase 7 of 13: Copying the third ocx control..."
Form2.ProgressBar1.Value = 7
FileCopy App.Path & "\RICHTX32.OCX", "C:\RICHTX32.OCX"
Form2.Label3.Caption = "Phase 8 of 13: Copying the fourth ocx control..."
Form2.ProgressBar1.Value = 8
FileCopy App.Path & "\COMDLG32.OCX", "C:\COMDLG32.OCX"
Form2.Label3.Caption = "Phase 9 of 13: Registering controls..."
Form2.ProgressBar1.Value = 9
    AppActivate Shell("C:\Starboyreg.bat")
    Form2.Label3.Caption = "Phase 10 of 13: Deleting temporary files..."
Form2.ProgressBar1.Value = 10
Kill "C:\MSCOMCT2.OCX"
    Form2.Label3.Caption = "Phase 11 of 13: Deleting temporary files..."
Form2.ProgressBar1.Value = 11
Kill "C:\TABCTL32.OCX"
    Form2.Label3.Caption = "Phase 12 of 13: Deleting temporary files..."
Form2.ProgressBar1.Value = 12
Kill "C:\RICHTX32.OCX"
    Form2.Label3.Caption = "Phase 13 of 13: Deleting temporary files..."
Form2.ProgressBar1.Value = 13
Kill "C:\COMDLG32.OCX"

If Check2.Value = 1 Then SaveDesktop
If Check1.Value = 1 Then Dialog.Show
If Check2.Value = 0 Then Stop

Form2.Label3.Caption = "All set! Your copy of Starboy Notepad has been successfully installed! In order to end the installation, close this window."
If Check2.Value = 1 Then SaveDesktop
Unload Form1
Unload FormWelcome
Unload Dialog
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame2.Visible = True
End Sub

Private Sub Command3_Click()
Form1.Hide
FormWelcome.Show
End Sub

Private Sub Command4_Click()
WebBrowser1.Visible = True
    Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir    'Remember the current active directory
    CommonDialog1.DialogTitle = "Select a directory" 'titlebar
    CommonDialog1.InitDir = App.Path 'start dir, might be "C:\" or so also
    CommonDialog1.FileName = "Select a Directory"  'Something in filenamebox
    CommonDialog1.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog1.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    CommonDialog1.CancelError = True 'allow escape key/cancel
    CommonDialog1.ShowSave   'show the dialog screen

    If Err <> 32755 Then    ' User didn't chose Cancel.
        Form1.Text1.Text = CurDir
    End If

    ChDir sTempDir  'restore path to what it was at entering
End Sub


Private Sub Command5_Click()
Unload FormWelcome
Unload Form1
End Sub

Private Sub Command6_Click()
Frame2.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command7_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command8_Click()
Unload FormWelcome
Unload Form1
End Sub

Private Sub Command9_Click()
Frame2.Visible = True
Frame3.Visible = False
End Sub

Private Sub Form_Load()
WebBrowser1.Visible = False
End Sub

Private Sub Text1_Change()
WebBrowser1.Navigate (Text1.Text)
End Sub
