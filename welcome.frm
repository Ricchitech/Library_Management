VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3915
   ClientLeft      =   4215
   ClientTop       =   2460
   ClientWidth     =   7680
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   Icon            =   "welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   2  'Custom
   Picture         =   "welcome.frx":2BC78
   ScaleHeight     =   3915
   ScaleWidth      =   7680
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   2
      Left            =   5040
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGEMENT"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 4
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Form9.Show
Unload Me
End If
End Sub
