VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Degration - beta"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Highscores"
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox High 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6600
      Left            =   1800
      Picture         =   "frmGame.frx":030A
      ScaleHeight     =   440
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   0
      Width           =   4500
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh List"
         Height          =   375
         Left            =   2520
         TabIndex        =   39
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   720
         TabIndex        =   60
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   720
         TabIndex        =   59
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   720
         TabIndex        =   58
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   720
         TabIndex        =   57
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   720
         TabIndex        =   56
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "15."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   55
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "14."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   54
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "13."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   53
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "12."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   52
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "11."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   51
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   49
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   48
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   47
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   46
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   45
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "10."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   44
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "9."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   43
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "8."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   42
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "7."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   41
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "6."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   40
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   38
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   37
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   36
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   35
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "HHH - 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Highscores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help!"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox buf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9600
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox blkM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6480
      Picture         =   "frmGame.frx":60E2C
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox blkS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6480
      Picture         =   "frmGame.frx":66C2E
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.ListBox lstHigh 
      Height          =   330
      Left            =   9240
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox block 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   9000
      Picture         =   "frmGame.frx":6CA30
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8760
      Top             =   1800
   End
   Begin VB.PictureBox mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9600
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Game"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6600
      Left            =   1800
      Picture         =   "frmGame.frx":6EFF2
      ScaleHeight     =   440
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      Begin VB.CommandButton cmdExHelp 
         Caption         =   "Exit"
         Height          =   375
         Left            =   1080
         TabIndex        =   33
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdN 
         Caption         =   "Next"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdAddHigh 
         Caption         =   "Add to Highscores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   15
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label lblHelp 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmGame.frx":CFB14
         Height          =   4695
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblBScore 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblRScore 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblTScore 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblGameover 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gameover"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblGameover 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gameover"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   225
         Width           =   3495
      End
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   96
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label lblSScore 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblSBlocks 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected:"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   8
      X2              =   96
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Label lblMoves 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Moves Left:"
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   112
      X2              =   112
      Y1              =   440
      Y2              =   0
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const tWidth As Long = 20

Dim Running As Boolean
Dim Paused As Boolean

Dim Score As Long
Dim tC As Long
Dim fC As Long
Dim Moves As Long
Dim ColYet As Boolean
Dim hX(50), hY(50), oX, oY

Private Function Collapse(B As PictureBox)
Dim X, Y, tX, tY, lT

If ColYet = False Then Exit Function

Score = Score + ((tC * 4) * 1.5)
lblScore.Caption = Score
tC = 0
lblMoves.Caption = Moves

For X = 14 To 0 Step -1
    For Y = 21 To 0 Step -1
        If B.Point(X, Y) = vbWhite Then
            B.PSet (X, Y), B.Point(X, Y - 1)
            B.PSet (X, Y - 1), vbWhite
        End If
    Next Y
Next X

Dim gDone As Boolean, rY As Long, lc As Boolean, gI As Long

gI = 1

OA:

If gI > 13 Then Exit Function

gDone = True

For X = 1 To 13
        If B.Point(X, 21) = vbWhite Then
            lc = False

            For rY = 0 To 21
                If B.Point(X, rY) <> vbWhite Then
                    lc = True
                End If
            Next rY

            If lc = False Then
                For Y = 0 To 21
                    B.PSet (X, Y), B.Point(X - 1, Y)
                    B.PSet (X - 1, Y), vbWhite
                    gDone = False
                Next Y
            End If
        End If
Next X

If gDone = False Then
    gI = gI + 1
    GoTo OA
End If
End Function

Function GameOver()
Dim X, Y, bS

tmrGame.Enabled = False
Running = False

Board.Cls

lblGameover(0).Visible = True
lblGameover(1).Visible = True
lblBScore.Visible = True
lblRScore.Visible = True
lblTScore.Visible = True
cmdAddHigh.Visible = True

lblRScore.Caption = "Score: " & Score

For X = 0 To 14
    For Y = 0 To 21
        If mask.Point(X, Y) <> vbWhite Then
            bS = bS + 1
        End If
    Next Y
Next X

lblBScore.Caption = "Blocks Left: " & bS
lblTScore.Caption = "Final Score: " & Score - bS
End Function

Function HiAmm() As Long
Dim i As Long

For i = 0 To 50
    If hX(i) > -1 Then HiAmm = HiAmm + 1
Next i
End Function

Function hSel(X, Y)
Dim mC

If buf.Point(X, Y) = vbWhite Then Exit Function

fC = fC + 1
mC = mask.Point(X, Y)
buf.PSet (X, Y), vbWhite
Hi X * tWidth, Y * tWidth

DoEvents

If buf.Point(X - 1, Y) = mC Then hSel X - 1, Y
If buf.Point(X + 1, Y) = mC Then hSel X + 1, Y
If buf.Point(X, Y - 1) = mC Then hSel X, Y - 1
If buf.Point(X, Y + 1) = mC Then hSel X, Y + 1
End Function

Function Sel(X, Y)
Dim tX, tY, mC

If mask.Point(X, Y) = vbWhite Then Exit Function

mC = mask.Point(X, Y)
mask.PSet (X, Y), vbWhite
DoExplo X * tWidth - 10, Y * tWidth - 10

tC = tC + 1

DoEvents

If mask.Point(X - 1, Y) = mC Then Sel X - 1, Y
If mask.Point(X + 1, Y) = mC Then Sel X + 1, Y
If mask.Point(X, Y - 1) = mC Then Sel X, Y - 1
If mask.Point(X, Y + 1) = mC Then Sel X, Y + 1
End Function

Private Sub Board_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tmrGame.Enabled = False Then Exit Sub

If mask.Point(Int(X / tWidth), Int(Y / tWidth)) = vbWhite Then Exit Sub

Moves = Moves - 1

If Moves < 0 Then
    Moves = 0
    lblMoves.Caption = Moves
    GameOver
    Exit Sub
End If

ClearHi

oX = -1
oY = -1
Board_MouseMove Button, Shift, X, Y

If HiAmm > 2 Then
    Sel Int(X / tWidth), Int(Y / tWidth)
Else
    Moves = Moves + 1
End If

ClearHi
oX = -1
oY = -1
Board_MouseMove Button, Shift, X, Y
End Sub

Function LoadScores()

End Function

Private Sub Board_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tmrGame.Enabled = False Then Exit Sub

If Int(X / tWidth) <> oX Or Int(Y / tWidth) <> oY And CheckHi(Int(X / tWidth), Int(Y / tWidth)) = False Then
oX = Int(X / tWidth)
oY = Int(Y / tWidth)

fC = 0
buf.PaintPicture mask.Image, 0, 0, , , , , , , vbSrcCopy
ClearHi
hSel oX, oY

lblSScore.Caption = "Score: " & ((fC * 4) * 1.5)
lblSBlocks.Caption = "Blocks: " & fC
End If
End Sub

Function CheckHi(X, Y) As Boolean
Dim i As Long

CheckHi = False

For i = 0 To 50
    If hX(i) = X And hY(i) = Y Then
        CheckHi = True
        Exit For
    End If
Next i
End Function

Private Sub cmdAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub cmdAddHigh_Click()
MsgBox "Sorry, This feature is not in the beta version", vbInformation, "BETA VERSION"
End Sub

Private Sub cmdExHelp_Click()
lblHelp.Visible = False
cmdN.Visible = False
cmdExHelp.Visible = False
tmrGame.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
tmrGame.Enabled = False
lblHelp.Visible = True
lblHelp.Caption = "How to Play               You select a group of 3 or more blocks by moving your mouse over them. Then you destroy them by clicking on them. The object of the game is too destroy the most ammount of blocks in the largest groups in 30 moves."
cmdN.Visible = True
cmdExHelp.Visible = True
End Sub

Private Sub cmdN_Click()
Select Case cmdN.Caption
Case "Next"
    lblHelp.Caption = "Your final score is calculated at the end by the score in the game deducted by ammount of blocks are left. If you score is high enough, you may be able to enter it into the highscores list"
    cmdN.Caption = "Prev"
Case "Prev"
    cmdHelp_Click
    cmdN.Caption = "Next"
End Select
End Sub

Private Sub cmdNew_Click()
Dim X, Y

High.Visible = False
Board.Visible = True

lblGameover(0).Visible = False
lblGameover(1).Visible = False
lblBScore.Visible = False
lblRScore.Visible = False
lblTScore.Visible = False
cmdAddHigh.Visible = False

For Y = 0 To 21
    For X = 1 To 13
        mask.PSet (X, Y), IIf(Int(Rnd * 2) = 0, IIf(Int(Rnd * 2) = 0, vbRed, vbBlue), IIf(Int(Rnd * 2) = 0, vbYellow, vbGreen))
    Next X
Next Y

ClearHi

Moves = 30
Score = 0
ColYet = True

Running = True
tmrGame.Enabled = True
End Sub

Function Hi(X, Y)
Dim i As Long

For i = 0 To 50
    If hX(i) = -1 Then
        hX(i) = X
        hY(i) = Y
        Exit For
    End If
Next i
End Function



Private Sub Command2_Click()
If Running = True Then
    Running = False
    Board.Visible = False
    High.Visible = True
Else
    Running = True
    Board.Visible = True
    High.Visible = False
End If
End Sub

Private Sub Form_Load()
High.Visible = True
Board.Visible = False
lblGameover(0).Visible = False
lblGameover(1).Visible = False
lblBScore.Visible = False
lblRScore.Visible = False
lblTScore.Visible = False
cmdAddHigh.Visible = False

Randomize
End Sub

Function Draw()
Dim X, Y, qY

If ColYet = False Then Exit Function

Board.Cls
Board.Line (0, 0)-(300, 440), RGB(245, 245, 245), BF

For X = 0 To 14
    For Y = 0 To 21
        If CheckHi(X * tWidth, Y * tWidth) = True Then
            qY = tWidth
        Else
            qY = 0
        End If
        Select Case mask.Point(X, Y)
        Case vbRed
            Board.PaintPicture block.Picture, X * tWidth, Y * tWidth, tWidth, tWidth, 0, qY, tWidth, tWidth, vbSrcCopy
        Case vbYellow
            Board.PaintPicture block.Picture, X * tWidth, Y * tWidth, tWidth, tWidth, tWidth, qY, tWidth, tWidth, vbSrcCopy
        Case vbBlue
            Board.PaintPicture block.Picture, X * tWidth, Y * tWidth, tWidth, tWidth, tWidth * 2, qY, tWidth, tWidth, vbSrcCopy
        Case vbGreen
            Board.PaintPicture block.Picture, X * tWidth, Y * tWidth, tWidth, tWidth, tWidth * 3, qY, tWidth, tWidth, vbSrcCopy
        Case vbMagenta
            Board.PaintPicture block.Picture, X * tWidth, Y * tWidth, tWidth, tWidth, tWidth * 4, 0, tWidth, tWidth, vbSrcCopy
        End Select
    Next Y
Next X
End Function

Function ClearHi()
Dim i As Long

For i = 0 To 50
    hX(i) = -1
    hY(i) = -1
Next i
End Function

Function DoExplo(X, Y)
Dim i As Long

For i = 0 To 4
    ColYet = False
    Board.PaintPicture blkM.Picture, X, Y, tWidth * 2, tWidth * 2, i * (tWidth * 2), 0, tWidth * 2, tWidth * 2, vbSrcAnd
    Board.PaintPicture blkS.Picture, X, Y, tWidth * 2, tWidth * 2, i * (tWidth * 2), 0, tWidth * 2, tWidth * 2, vbSrcInvert
    DoEvents
Next i

ColYet = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Running = False
tmrGame.Enabled = False
End Sub

Private Sub tmrGame_Timer()
If Running = False Then Exit Sub
Collapse mask
Draw
End Sub
