VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest For Glory Importer ’95"
   ClientHeight    =   5055
   ClientLeft      =   3060
   ClientTop       =   2055
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "formMain"
   MaxButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDragonFire 
      Caption         =   "Importing Into Quest For Glory V"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      ToolTipText     =   "Dragon Fire"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CheckBox chkPerfect 
      Caption         =   "&Perfect Characters"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Every Skill At Max"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox chkSillyClowns 
      Caption         =   "&Silly Clowns"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Check It And See!"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtGame 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Quest For Glory I"
      ToolTipText     =   "Importing From"
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox txtClass 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Fighter"
      ToolTipText     =   "Character Type"
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtCharName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   360
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "Unknown Hero"
      ToolTipText     =   "Character Name"
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H80000005&
      Caption         =   "Pre&view"
      Height          =   495
      Left            =   2160
      TabIndex        =   17
      ToolTipText     =   "Look at the Character"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame frmGame 
      Caption         =   "Importing From:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Chose A Game"
      Top             =   3000
      Width           =   4215
      Begin VB.OptionButton optGlory 
         Caption         =   "Quest For Glory IV"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   14
         ToolTipText     =   "Shadows Of Darkness"
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optGlory 
         Caption         =   "Quest For Glory III"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Wages Of War"
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optGlory 
         Caption         =   "Quest For Glory II"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   12
         ToolTipText     =   "Trial By Fire"
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optGlory 
         Caption         =   "Quest For Glory I"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "So You Want To Be A Hero"
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdExitButton 
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Exit Program"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveBox 
      BackColor       =   &H80000005&
      Caption         =   "C&reate Character"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Save Character"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame frmCharClass 
      Caption         =   "Character Class:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Chose A Character"
      Top             =   1920
      Width           =   4215
      Begin VB.OptionButton optClass 
         Caption         =   "Paladin"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Thief"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Magic User"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optClass 
         Caption         =   "Fighter"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "sav"
      DialogTitle     =   "Save Character"
      Filter          =   "QFG Chatacter Files (*.sav)|*.sav|All Files (*.*)|*.*|"
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quest For Glory Importer ’95"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v0.00.0000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   20
      Top             =   360
      Width           =   810
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePreview 
         Caption         =   "&Preview"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Character As..."
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePerfect 
         Caption         =   "Per&fect Characters"
      End
      Begin VB.Menu mnuFileSillyClowns 
         Caption         =   "Silly &Clowns"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpRegister 
         Caption         =   "&Register!"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About QFG Importer '95"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Glory1 = 1, Glory2 = 2, Glory3 = 3, Glory4 = 4
Const Fighter = 1, Wizard = 2, Thief = 3, Paladin = 4
Dim WhatsThePicture, InitDir, Path
Dim FighterFile(1 To 4) As String
Dim WizardFile(1 To 4) As String
Dim ThiefFile(1 To 4) As String
Dim PaladinFile(2 To 4) As String
Dim FighterPerfect(1 To 4) As String
Dim WizardPerfect(1 To 4) As String
Dim ThiefPerfect(1 To 4) As String
Dim PaladinPerfect(2 To 4) As String
Dim SaveFileName As String
Dim X
Private Sub MakeFileName()
    If chkDragonFire.Value = vbChecked Then
        dlgCommonDialog.FileName = "Quest For Glory "
    Else
        dlgCommonDialog.FileName = "Qg"
    End If
    For X = 1 To 4
        If optGlory(X).Value = True Then
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & X
        End If
    Next X
    If chkDragonFire.Value = vbChecked Then
        If chkPerfect.Value = vbChecked Then
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & " Perfect"
        End If
    End If
    If optClass(Fighter).Value = True Then
        If chkDragonFire.Value = vbChecked Then
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & " Fighter"
        Else
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & "Fghtr"
        End If
    ElseIf optClass(Wizard).Value = True Then
        If chkDragonFire.Value = vbChecked Then
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & " Wizard"
        Else
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & "Wizrd"
        End If
    ElseIf optClass(Thief).Value = True Then
        If chkDragonFire.Value = vbChecked Then
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & " Thief"
        Else
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & "Thief"
        End If
    ElseIf optClass(Paladin).Value = True Then
        If chkDragonFire.Value = vbChecked Then
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & " Paladin"
        Else
            dlgCommonDialog.FileName = dlgCommonDialog.FileName & "Paldn"
        End If
    End If
    dlgCommonDialog.FileName = dlgCommonDialog.FileName + ".sav"
End Sub
Private Sub LoadCharacters()
'Quest For Glory I
'Version 3.0 Characters
FighterFile(Glory1) = "53565637 460 460 460 460 4 4 46060607746262626262626262626262c2c2c2c2c5553414a 9 12c5c"
WizardFile(Glory1) = "52575737 561 561 561 5 56161616161 510395b61 561 561 561 5616060606060191f 66e2d25 878"
ThiefFile(Glory1) = "51545434 a6e a6e a6e a a6e a6e a6e6e795c3939393939393939393932323232324b4d205e1d153848"
'No Known Problems/Inconsistancies
'Version 1.0 Perfect Characters
FighterPerfect(Glory1) = "53565637 86c 86c 86c 86c 86c 86c 86c13337a284c284c284c284c2822222222225b5d1a3e7d755828"
WizardPerfect(Glory1) = "52575737 96d 96d 96d 96d 96d 96d 96d1a3b7d197d197d197d197d1912121212126b6d5066252d 070"
ThiefPerfect(Glory1) = "51545434 a6e a6e a6e a6e a6e a6e a6e795c345034503450345034505b5b5b5b5b2224241e5d5578 8"

'Quest For Glory II
'Version 3.0 Characters
FighterFile(Glory2) = "535353199c669c669c669c669c9c9c6666669c66bb26747474747474747474747474747e1a1a1a7edee034 52abaa3 0"
WizardFile(Glory2) = "52525212966c966c966c96966c6c6c6c6c966c96bd18b6669c669c669c669c669c669c9cf89c9cf858663c7d52c2db78"
ThiefFile(Glory2) = "5151516eea10ea10ea10eaea10ea10ea1010ea10c93dbbbbbbbbbbbbbbbbbbbbbbbbbbb1d5d5d5b1112fd6ab8414 dae"
PaladinFile(Glory2) = "50505091 ef4 ef4 ef4 ef4 e e ef4f4f4 ef49a9cd4d4d4d4d4d4d4d4d4d4d4d4d4debababade7e403829 6968f2c"
'No Known Problems/Inconsistancies
'Version 1.0 Perfect Characters
FighterPerfect(Glory2) = "535353199c669c669c669c669c669c669c669c66ca3ed62cd62cd62cd62cd62cd62cd6dcb8b8b8dc7c424b133cacb516"
WizardPerfect(Glory2) = "52525212966c966c966c966c966c966c966c966cc034dc26dc26dc26dc26dc26dc26dcd6b2d6d6b2122c2597b8283192"
ThiefPerfect(Glory2) = "5151516eea10ea10ea10ea10ea10ea10ea10ea10f8 cd228d228d228d228d228d228d2d8bcbcbcd878467a3718889132"

'I just have the fighter in place temporarily...
PaladinPerfect(Glory2) = FighterPerfect(Glory2)

'Quest For Glory III
'Version 2.0 Characters
FighterFile(Glory3) = " 053 052 019 f33121b f33121b f33121b f33121b f33 f33 f33121b121b121b f33121b147 91444314443144431444314443144431444314443144431444314443144431444314443144431444314443144431444314437144371441114411144351444c1433613137ffdf34ffdf23ffde62ffde1effde2c"
PaladinFile(Glory3) = " 050 051 1 7 1 9 315 1 9 315 1 9 315 1 9 315 1 9 1 9 1 9 315 315 1 9 315 1 9ff513dff4e63ff4e63ff511bff511bff511bff511bff511bff511bff511bff511bff511bff511bff511bff511bff511bff511bff4e63ff4f 5ff4f 5ff4e5aff4e5aff4d58ff4d58ff4d58ff4d61ff4f37ff43441315613129130601325c13056"
ThiefFile(Glory3) = " 051 050 1 1 44b 22f 44b 22f 44b 22f 44b 44b 22f 44b 22f 44b 22f 22f 44b 22f2d392b152b182b182b182b182b182b182b182b182b182b182b182b182b182b182b182b182b182b182b232b222a222a222b1d2b24294a41 e d62 d2d d18 f c d 6"
'Version 1.0 Characters
WizardFile(Glory3) = " 052 053 0 9 133 42b 133 42b 133 42b 133 133 42b 42b 42b 42b 42b 133 42b 133f949f733f72ef942f72ef638f938f638f938f638f938f638f938f638f938f638f647f92ff92ff92ff92ef92ff84af834f932f92bfa2911f2b4f4e4f415030511050 6"
'Paladin has 17000+ Open skill when imported
'All Characters have Magic -- They shouldn't
'Version 1.0 Perfect Characters
FighterPerfect(Glory3) = ""
WizardPerfect(Glory3) = ""
ThiefPerfect(Glory3) = ""
PaladinPerfect(Glory3) = ""

'Quest For Glory IV
'Version 1.0 Characters
FighterFile(Glory4) = " 053 050 05f 04f 44f 04f 44f 04f 44f 04f 44f 04f 04f 04f 44f 04f 04f 44f 04f 04fffffff5b41ffffff5819ffffff5c 4ffffff5c 4ffffff58 effffff5a62ffffff58 effffff5a62ffffff58 effffff5a62ffffff58 effffff5a62ffffff58 effffff5a62ffffff58 effffff5a62ffffff5a62ffffff58 effffff58 effffff58 effffff58 effffff58 effffff58 effffff58 effffff58 effffff5a62ffffff5a62ffffff5a62ffffff5a62ffffff5a62ffffff5a62ffffff5a62ffffff5b17ffffff5b591191d137321371d138141361813746"
WizardFile(Glory4) = " 052 050 052 042 442 042 442 042 442 042 042 442 442 442 442 442 042 442 042 042ffffffbc59ffffffbb4dffffffbc59ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbb41ffffffbd 9ffffffbc11ffffffba49ffffffba49ffffffba49ffffffba49ffffffba49ffffffba49ffffffba49ffffffba49ffffffba50ffffffb94affffffba5cfffffffb2afffffffb3dfffffffa44fffffffa 8fffffffb16"
ThiefFile(Glory4) = " 051 052 04b 117 45b 117 45b 117 45b 117 117 45b 117 45b 117 45b 45b 117 45b 1175d a5c2f5e445e443a493d 93a493d 93a493d 93a493d 93a493d 93a493d 93d 93a493a493a493a493a493a493a493a493d 93d 93d 93d 93d 93d 93d 93c543c a11f2a1051e10531105541063410626"
PaladinFile(Glory4) = " 050 052 0 0 010 354 010 354 010 354 010 354 010 010 010 354 010 354 010 354 354fffffee22afffffee62afffffee22afffffee62affffff95 fffffff91 fffffff95 fffffff91 fffffff95 fffffff91 fffffff95 fffffff91 fffffff95 fffffff91 fffffff95 fffffff91 fffffff91 fffffff95 fffffff91 fffffff91 fffffff95 fffffff91 fffffff95 fffffff95 fffffff91 fffffff935bffffff935bffffff935bffffff935bffffff935bffffff935bffffff935bffffff94 effffff94585c509f629f2d9f 8a0609e5a"
'All Characters have Magic -- They shouldn't
'Version 1.0 Perfect Characters
FighterPerfect(Glory4) = ""
WizardPerfect(Glory4) = ""
ThiefPerfect(Glory4) = ""
PaladinPerfect(Glory4) = ""


'Backup Characters/Old Characters
'QfG I
'Version 1.0 Characters
'FighterFile(Glory1) = "53515131 266 266 266 266 2 2 266666673561616161616161616161616161616166f69 f74373f1262"
'WizardFile(Glory1) = "52505030 064 064 064 0 06464646464 015 e492d492d492d492d492d2c2c2c2c2c55532f2e6d654838"
'ThiefFile(Glory1) = "51535333 b6f b6f b6f b b6f b6f b6f6f7a6e2424242424242424242425252525255c5a1a70333b1666"
'Version 2.0 Characters
'FighterFile(Glory1) = "53565635 a6e a6e a6e a6e a a a6e6e6e125a4a4a4a4a4a4a4a4a4a4a4040404040393f c286b634e3e"
'WizardFile(Glory1) = "52575737 86c 86c 86c 8 86c6c6c6c6c 8 941513551355135513551353f3f3f3f3f4640 9387b735e2e"
'ThiefFile(Glory1) = "51545434 b6f b6f b6f b b6f b6f b6f6f6a77747410741074107410747e7e7e7e7e 7 1 8 e4d456818"
'QfG II
'Version 1.0 Characters
'FighterFile(Glory2) = "535353ce51ab51ab51ab51ab515151ababab51abaabfe7e7e7e7e7e7e7e7e7e7e7e7e7e783e7c0a4 43a 5537cecf556"
'WizardFile(Glory2) = "525252252ad02ad02ad02a2ad0d0d0d0d02ad02add578837cd37cd37cd37cd37cd37cdcca8cceb8f2f1118557aeaf350"
'ThiefFile(Glory2) = "515151222dd72dd72dd72d2dd72dd72dd7d72dd774d0 7 7 7 7 7 7 7 7 7 7 7 7 7 561 5 561c1ff60ddf2627bd8"
'PaladinFile(Glory2) = "505050b12ed42ed42ed42ed42e2e2ed4d4d42ed43f c444444444444444444444444444420444c2888b64bf3dc4c55f6"
'Version 2.0 Characters
'FighterFile(Glory2) = "535353199c669c669c669c669c9c9c6666669c6661197c9b619b619b619b619b619b616b f6b6b faf917c456afae340"
'WizardFile(Glory2) = "52525212966c966c966c96966c6c6c6c6c966ca4926b16cf35cf35cf35cf35cf35cf353551353551f1cffb97b8283192"
'ThiefFile(Glory2) = "5151516eea10ea10ea10eaea10d822d82222d822df2bbbbb41bb41bb41bb41bb41bb414b2f4b4b2f8fb1b223 c9c8526"
'PaladinFile(Glory2) = "50505091 ef4 ef4 ef4 ef4 e e ef4f4f4 ef4bc5fad2bd12bd12bd12bd12bd12bd1dbbfdbdbbf1f21f9436cfce546"
'QfG III
'Version 1.0 Characters
'FighterFile(Glory3) = " 053 052 019 f33121b f33121b f33121b f33121b f33 f33 f33121b121b121b f33121b147 a1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1444b1443f1443f14419144191443d144441433e131 6ffde48ffde57ffdf effdd12ffdf38"
'PaladinFile(Glory3) = " 050 051 062 110 32c 110 32c 110 32c 110 32c 110 110 110 32c 32c 331 1 d 33112b40129 712863129601296012960129601296012960129601296012960129601296012960129601296012a 6129601296012a e12a e129 c129 c129 c129 512a4f1403c14221359133c13 01332"
'ThiefFile(Glory3) = " 051 053 057 459 161 459 161 459 161 459 459 161 459 161 459 24b 24b 03f 24b134251364113644136441364413644136441364413644136441364413644136441364413644136441364413644136441364413642136431355713657135571355e1373012043ffe0 8ffdf5bffe062ffe13affe02c"
'QfG IV
End Sub
Private Sub ChangePreview()
    frmPreview.lblName.BorderStyle = False
    frmPreview.Caption = "Character Preview - " & txtClass
    If txtGame = "Quest For Glory I" Then
        If chkSillyClowns.Value = False Then
            frmPreview.lblName.Move 2140, 325
            frmPreview.lblName.ForeColor = vbWindowText '&H80000008
            frmPreview.cmdOK.Move 3960, 2400
            WhatsThePicture = 1
        Else
            frmPreview.lblName.Move 840, 68
            frmPreview.lblName.ForeColor = vbBlue '&HFF0000
            frmPreview.cmdOK.Move 3120, 2160
            WhatsThePicture = 0
        End If
    ElseIf txtGame = "Quest For Glory II" Then
        frmPreview.lblName.Move 840, 68
        frmPreview.lblName.ForeColor = vbBlue '&HFF0000
        frmPreview.cmdOK.Move 3120, 2160
        WhatsThePicture = 2
    ElseIf txtGame = "Quest For Glory III" Then
        frmPreview.lblName.Move 1544, 65
        frmPreview.lblName.ForeColor = vbWindowText '&H80000008
        frmPreview.cmdOK.Move 2040, 2400
        WhatsThePicture = 3
    ElseIf txtGame = "Quest For Glory IV" Then
        frmPreview.lblName.Move 1800, 2520
        frmPreview.lblName.ForeColor = vbWindowText '&H80000008
        frmPreview.cmdOK.Move 120, 2400
        WhatsThePicture = 4
    End If

'*****************************
'MsgBox txtClass
'*****************************

    If chkPerfect.Value = vbChecked Then
        If txtClass.Text = "Fighter" Then
            frmPreview.Picture = frmPictures.picPerfectFighter(WhatsThePicture)
        ElseIf txtClass.Text = "Wizard" Or txtClass.Text = "Magic User" Then
            frmPreview.Picture = frmPictures.picPerfectWizard(WhatsThePicture)
        ElseIf txtClass.Text = "Thief" Then
            frmPreview.Picture = frmPictures.picPerfectThief(WhatsThePicture)
        ElseIf txtClass.Text = "Paladin" Then
            frmPreview.Picture = frmPictures.picPerfectPaladin(WhatsThePicture)
        End If
    Else
        If txtClass.Text = "Fighter" Then
            frmPreview.Picture = frmPictures.picFighter(WhatsThePicture)
        ElseIf txtClass.Text = "Wizard" Or txtClass.Text = "Magic User" Then
            frmPreview.Picture = frmPictures.picWizard(WhatsThePicture)
        ElseIf txtClass.Text = "Thief" Then
            frmPreview.Picture = frmPictures.picThief(WhatsThePicture)
        ElseIf txtClass.Text = "Paladin" Then
            frmPreview.Picture = frmPictures.picPaladin(WhatsThePicture)
        End If
    End If
End Sub

Private Sub chkPerfect_Click()
    If chkPerfect.Value = vbChecked Then
        optGlory(Glory3).Enabled = False
        optGlory(Glory4).Enabled = False
        If optGlory(Glory3).Value = True Or optGlory(Glory4).Value = True Then
            optGlory(Glory1).Value = True
        End If
        mnuFilePerfect.Checked = True
    Else
        optGlory(Glory3).Enabled = True
        optGlory(Glory4).Enabled = True
        mnuFilePerfect.Checked = False
    End If
    mnuFilePerfect.Checked = chkPerfect.Value
    Call ChangePreview
End Sub

Private Sub chkSillyClowns_Click()
    mnuFileSillyClowns.Checked = chkSillyClowns.Value
    If chkSillyClowns = vbChecked Then
        frmSplash.imgLogo = frmPictures.picWaffle
    Else
        frmSplash.imgLogo = frmPictures.picQFG
    End If
    Call ChangePreview
End Sub

Private Sub cmdExitButton_Click()
    Unload Me
    Unload frmPreview
    Unload frmPictures
    Unload frmSplash
    End
End Sub

Private Sub cmdSaveBox_Click()
    Call MakeFileName
    InitDir = GetSetting(App.CompanyName & "\" & App.EXEName, "Settings", "StartIn", App.Path)
'MsgBox InitDir
    dlgCommonDialog.InitDir = InitDir
On Error GoTo ErrorHandle
    dlgCommonDialog.ShowSave
ErrorHandle:
    If Err = cdlCancel Then   '32755 = Cancel
        'MsgBox "You Canceled!"
    ElseIf Err = 0 Then  '0 = Save
        Call LoadCharacters
        SaveFileName = ""
        Dim ThisGame As Integer, ThisClass As Integer
        For X = 1 To 4
          If optGlory(X).Value = True Then
            ThisGame = X
          End If
        Next
        For X = 1 To 4
          If optClass(X).Value = True Then
            ThisClass = X
          End If
        Next
        'MsgBox "Game = " & ThisGame & vbCrLf & "Class = " & ThisClass
        
        If ThisGame = Glory3 Then
            'Print #1, " glory3.sav"
            SaveFileName = SaveFileName & " glory3.sav " & vbLf
        ElseIf ThisGame = Glory4 Then
            'Print #1, " glory4.sav "
            SaveFileName = SaveFileName & " glory4.sav" & vbLf
        End If
        'Print #1, frmMain.txtCharName.Text
        SaveFileName = SaveFileName & frmMain.txtCharName.Text & vbLf

        If ThisClass = Fighter Then
            If chkPerfect.Value = vbChecked Then
                SaveFileName = SaveFileName & FighterPerfect(ThisGame) & vbLf
            Else
                SaveFileName = SaveFileName & FighterFile(ThisGame) & vbLf
            End If
        ElseIf ThisClass = Wizard Then
            If chkPerfect.Value = vbChecked Then
                SaveFileName = SaveFileName & WizardPerfect(ThisGame) & vbLf
            Else
                SaveFileName = SaveFileName & WizardFile(ThisGame) & vbLf
            End If
        ElseIf ThisClass = Thief Then
            If chkPerfect.Value = vbChecked Then
                SaveFileName = SaveFileName & ThiefPerfect(ThisGame) & vbLf
            Else
                SaveFileName = SaveFileName & ThiefFile(ThisGame) & vbLf
            End If
        ElseIf ThisClass = Paladin Then
            If chkPerfect.Value = vbChecked Then
                SaveFileName = SaveFileName & PaladinPerfect(ThisGame) & vbLf
            Else
                SaveFileName = SaveFileName & PaladinFile(ThisGame) & vbLf
            End If
        End If
        
        'Print #1, "This character was created using the QFG Importer '95"
        SaveFileName = SaveFileName & vbCrLf & "This character was created using the QFG Importer ’95 v" & App.Major & "." & Right$("00" & App.Minor, 2)
        If Val(App.Revision) <> 0 Then
            SaveFileName = SaveFileName & "." & Right$("0000" & App.Revision, 4)
        End If
        Open dlgCommonDialog.FileName For Output As #1
            Print #1, SaveFileName
        Close
        
        MsgBox "Character Creation in “" & dlgCommonDialog.FileName & "” Successful."
        If ThisGame = Glory3 Then
            MsgBox "NOTE: Some versions of Quest For Glory IV might not import this character." & vbCrLf & "  If this happens to you, simply download the QG4PAT file from www.sierra.com"
            If ThisClass = Fighter Then
                MsgBox "NOTE: Some versions of Quest For Glory IV might think that this character is a Wizard." & vbCrLf & "  If this happens, simply choose Fighter and continue."
            End If
        End If
    Else
        MsgBox "Error " & Err & ".  " & Err.Description, , "QfG Importer ’95 Error"
    End If
    Dim i
    For i = Len(Path) To 1 Step -1
        If Mid(Path, i, 1) <> "\" Then
            If Path <> "" Then
                Path = Left(Path, i - 1)
            End If
        Else
            Exit For
        End If
    Next i
    SaveSetting App.CompanyName & "\" & App.EXEName, "Settings", "StartIn", Path
End Sub

Private Sub cmdPreview_Click()
    Call ChangePreview
    frmPreview.Show
    mnuFilePreview.Checked = True
    frmPreview.lblName.Caption = txtCharName.Text
End Sub

Private Sub Form_GotFocus()
    If frmSplash.Visible = True Then
        frmSplash.SetFocus
    End If
End Sub

Private Sub Form_Load()
    lblVersion = "v" & App.Major & "." & Right$("00" & App.Minor, 2)
    lblVersion.ToolTipText = "v" & App.Major & "." & Right$("00" & App.Minor, 2) & "  Build " & Right$("0000" & App.Revision, 4)
'    If Val(App.Revision) = 0 Then
'        lblVersion.ToolTipText = "v" & App.Major & "." & Right$("00" & App.Minor, 2)
'    Else
'        lblVersion = "v" & App.Major & "." & Right$("00" & App.Minor, 2) & "  Build " & Right$("0000" & App.Revision, 4)
'    End If
    dlgCommonDialog.Flags = cdlOFNExtensionDifferent + cdlOFNLongNames _
                            + cdlOFNOverwritePrompt + cdlOFNHideReadOnly _
                            + cdlOFNExplorer
                            '1038
    'MsgBox dlgCommonDialog.Flags
    frmPreview.lblName.BorderStyle = vbTransparent '0
    frmPreview.lblName.Move 2140, 325
    frmPreview.lblName.ForeColor = vbWindowText '&H80000008
    frmPreview.cmdOK.Move 3960, 2400
    frmPreview.Picture = frmPictures.picFighter(Glory1).Picture
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState = vbMinimized Then '1
        frmPreview.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPictures
    Unload frmPreview
    Unload frmMain
'    MsgBox "Good-Bye!", vbExclamation, "QfG Importer '95"
    End
End Sub

Private Sub mnuHelpAbout_Click()
    frmSplash.Show
    frmMain.Enabled = False
    frmSplash.SetFocus
'    frmSplash.lblWarning.Width = 369
End Sub

Private Sub mnuFileExit_Click()
    Call cmdExitButton_Click
End Sub

Private Sub mnuFilePerfect_Click()
    If chkPerfect.Value = vbChecked Then '1
        chkPerfect.Value = vbUnchecked '0
    Else
        chkPerfect.Value = vbChecked '1
    End If
    Call chkPerfect_Click
End Sub

Private Sub mnuFilePreview_Click()
    If mnuFilePreview.Checked = False Then
        Call cmdPreview_Click
        mnuFilePreview.Checked = True
    Else
        frmPreview.Hide
        mnuFilePreview.Checked = False
    End If
End Sub

Private Sub mnuFileSave_Click()
    Call cmdSaveBox_Click
End Sub

Private Sub mnuFileSillyClowns_Click()
    If chkSillyClowns.Value = vbChecked Then '1
        chkSillyClowns.Value = vbUnchecked '0
    Else
        chkSillyClowns.Value = vbChecked '1
    End If
    Call chkSillyClowns_Click
End Sub

Private Sub mnuHelpRegister_Click()
    frmSplash.txtUserName.Visible = True
    frmSplash.lblUserName.Visible = True
    frmSplash.txtUserNumber.Visible = True
    frmSplash.lblUserNumber.Visible = True
    frmSplash.cmdSysInfo.Visible = False
    frmSplash.cmdOK.Default = True
    Call mnuHelpAbout_Click
    frmSplash.txtUserName.SetFocus
End Sub

Private Sub optClass_Click(Index As Integer)
    txtClass = optClass(Index).Caption
    Call ChangePreview
End Sub

Private Sub optGlory_Click(Index As Integer)
    txtGame = optGlory(Index).Caption
    If Index = Glory1 Then
        optClass(Paladin).Enabled = False
        optClass(Wizard).Caption = "Magic User"
        If optClass(Wizard).Value = True Then
            txtClass.Text = "Magic User"
        End If
        If optClass(Paladin).Value = True Then
            optClass(Fighter).Value = True
        End If
    Else
        optClass(Paladin).Enabled = True
        optClass(Wizard).Caption = "Wizard"
        If optClass(Wizard).Value = True Then
            txtClass.Text = "Wizard"
        End If
    End If
    Call ChangePreview
End Sub

Private Sub txtCharName_Change()
    frmPreview.lblName.Caption = txtCharName.Text
End Sub

Private Sub txtCharName_GotFocus()
    txtCharName.SelStart = 0
    txtCharName.SelLength = Len(txtCharName)
End Sub

Private Sub txtCharName_LostFocus()
    txtCharName.SelLength = 0
    If txtCharName.Text = "" Then
        txtCharName.Text = "Unknown Hero"
    End If
End Sub
