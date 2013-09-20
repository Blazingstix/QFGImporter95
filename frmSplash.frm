VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtUserNumber 
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdSecret 
      Caption         =   "Show me the Secrets!!!"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   8
      Top             =   3720
      Width           =   1485
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   9
      Top             =   4080
      Width           =   1485
   End
   Begin VB.Label lblUserNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Number:"
      Height          =   210
      Left            =   1800
      TabIndex        =   16
      Top             =   2925
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   210
      Left            =   1800
      TabIndex        =   15
      Top             =   2325
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblIrwinCo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IrwinCo."
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblSierra 
      BackStyle       =   0  'Transparent
      Caption         =   "All images are copyright of Sierra On-Line..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   150
      TabIndex        =   10
      Tag             =   "LicenseTo"
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   8
      X2              =   489
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   489
      X2              =   489
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   8
      X2              =   489
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   8
      X2              =   8
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Image imgLogo 
      Height          =   1290
      Left            =   240
      Picture         =   "frmSplash.frx":1A7EB
      Top             =   540
      Width           =   1290
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1998"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Tag             =   "Copyright"
      Top             =   2460
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Company IrwinCo."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Tag             =   "Company"
      Top             =   2700
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Warning"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   795
      Left            =   300
      TabIndex        =   5
      Tag             =   "Warning"
      Top             =   3660
      Width           =   6855
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.00.0000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4650
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 95/98"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4245
      TabIndex        =   3
      Tag             =   "Platform"
      Top             =   2520
      Width           =   2835
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IrwinCo. Presents..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1680
      TabIndex        =   2
      Tag             =   "CompanyProduct"
      Top             =   420
      Width           =   4035
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product: QFG ’95"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   1680
      TabIndex        =   1
      Tag             =   "Product"
      Top             =   900
      Width           =   5265
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed To: Charles Irwin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "LicenseTo"
      Top             =   150
      Width           =   6855
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsButtonDown As Boolean
Dim FirstX, FirstY
Dim SecretsRevealed As Boolean
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub CheckRegistration()
    If txtUserName <> "" And txtUserNumber <> "" Then
        MsgBox "That Registration Number was Invalid.", , "Registration"
    End If
    txtUserName = ""
    txtUserNumber = ""
End Sub
Private Sub cmdSecret_Click()
    frmMain.chkSillyClowns.Visible = True
    frmMain.mnuFileSillyClowns.Visible = True
    SecretsRevealed = True
    Call cmdOK_Click
    lblProductName.Top = 60
    lblCompanyProduct.Top = 28
    lblIrwinCo.Move 16, 160
    cmdSecret.Visible = False
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
'    Unload Me
    'Me.Hide
    
    If txtUserName.Visible = True And txtUserNumber.Visible = True Then
        Me.Hide
        Call CheckRegistration
    Else
        Me.Hide
    End If
    
    txtUserName.Visible = False
    lblUserName.Visible = False
    txtUserNumber.Visible = False
    lblUserNumber.Visible = False
    cmdSysInfo.Visible = True
    cmdOK.Default = False
    frmMain.Enabled = True
    frmMain.SetFocus
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Load()
    If Val(App.Revision) = 0 Then
        lblVersion = "Version " & App.Major & "." & Right$("00" & App.Minor, 2)
    Else
        lblVersion = "Version " & App.Major & "." & Right$("00" & App.Minor, 2) & "  Build " & Right$("0000" & App.Revision, 4)
    End If
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCopyright = App.LegalCopyright
    lblWarning = App.Comments & "!"
    lblCompany = App.CompanyName
    IsButtonDown = False
End Sub

Private Sub lblIrwinCo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SecretsRevealed = False Then
        If Shift = 2 And Button = 1 And txtUserName.Visible = False Then
            IsButtonDown = True
            FirstX = X / Screen.TwipsPerPixelX
            FirstY = Y / Screen.TwipsPerPixelY
        End If
    End If
End Sub
    
Private Sub lblIrwinCo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 2 And Button = 1 Then
        If IsButtonDown = True Then
            X = X / Screen.TwipsPerPixelX
            Y = Y / Screen.TwipsPerPixelY
            
            'MsgBox "FirstX = " & FirstX & vbCrLf & "X = " & X & vbCrLf & "X - FirstX = " & X - FirstX
            'MsgBox "FirstY = " & FirstY & vbCrLf & "Y = " & Y & vbCrLf & "Y - FirstY = " & Y - FirstY
            'lblIrwinCo.Left = lblIrwinCo.Left + ((X / Screen.TwipsPerPixelX) - (FirstX / Screen.TwipsPerPixelX))
            'lblIrwinCo.Top = lblIrwinCo.Top + ((Y / Screen.TwipsPerPixelY) - (FirstY / Screen.TwipsPerPixelY))
            lblIrwinCo.Left = lblIrwinCo.Left + (X - FirstX)
            lblIrwinCo.Top = lblIrwinCo.Top + (Y - FirstY)
            
            If SecretsRevealed = False Then
                If lblProductName.Top > -55 Then
                    If lblIrwinCo.Top <= 110 Then '(lblProductName.Height - lblProductName.Top) Then
                        lblCompanyProduct.Top = lblCompanyProduct.Top + (Y - FirstY)
                        lblProductName.Top = lblProductName.Top + (Y - FirstY)
                        If lblProductName.Top > 60 Then
                            lblProductName.Top = 60
                            lblCompanyProduct.Top = 28
                        End If
                    End If
                Else
                    cmdSecret.Visible = True
                End If
            End If
        End If
    End If
End Sub

Private Sub lblIrwinCo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsButtonDown = False
End Sub
