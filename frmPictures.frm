VERSION 5.00
Begin VB.Form frmPictures 
   Caption         =   "This Form has the Pictures on it"
   ClientHeight    =   5400
   ClientLeft      =   1140
   ClientTop       =   1725
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmPictures.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picQFG 
      Height          =   975
      Left            =   1320
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   37
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox picWaffle 
      Height          =   975
      Left            =   120
      Picture         =   "frmPictures.frx":000C
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   36
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox picPerfectPaladin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4800
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   35
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectPaladin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   3600
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   34
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectPaladin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":07B0
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":1B2A
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "frmPictures.frx":2ECC
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmPictures.frx":D7F5
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   27
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":E9D8
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   25
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "frmPictures.frx":FDE3
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmPictures.frx":1A7C7
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":1BA27
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "frmPictures.frx":1CDA5
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPerfectFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmPictures.frx":2781C
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmPictures.frx":28AA3
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmPictures.frx":29C66
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmPictures.frx":2AEA2
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPaladin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      Picture         =   "frmPictures.frx":2C114
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      Picture         =   "frmPictures.frx":35578
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      Picture         =   "frmPictures.frx":3E9E6
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4800
      Picture         =   "frmPictures.frx":47D4B
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPaladin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":511D0
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":5251A
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":5385C
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   2400
      Picture         =   "frmPictures.frx":54BDC
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      Picture         =   "frmPictures.frx":55FBA
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPaladin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      Picture         =   "frmPictures.frx":5DC0E
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      Picture         =   "frmPictures.frx":6587B
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3600
      Picture         =   "frmPictures.frx":6D4E4
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picThief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "frmPictures.frx":75189
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "frmPictures.frx":7FAD9
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picFighter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1200
      Picture         =   "frmPictures.frx":8A4C3
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line lneTheLineThatSeperates 
      Index           =   1
      X1              =   120
      X2              =   5880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line lneTheLineThatSeperates 
      Index           =   0
      X1              =   120
      X2              =   5880
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

