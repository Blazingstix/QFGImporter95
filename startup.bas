Attribute VB_Name = "Startup"
Option Explicit

Sub Main()
    
    Dim TheTime, CanIGo As Boolean, IsThereRegistration As Boolean, IsRegistered As Boolean
    IsThereRegistration = False
    IsRegistered = True
    
    TheTime = Now
    
    Load frmSplash
    
    If IsThereRegistration = True Then
        frmSplash.lblLicenseTo = "UNREGISTERED!!!"
    Else
        frmSplash.lblLicenseTo = ""
    End If
    frmSplash.lblIrwinCo.Visible = False
    frmSplash.cmdOK.Visible = False
    frmSplash.cmdSysInfo.Visible = False
    
    frmSplash.Show
    frmSplash.Refresh
    
    If IsThereRegistration = True Then
        If frmSplash.lblLicenseTo = "UNREGISTERED!!!" Then
            IsRegistered = False
        Else
            IsRegistered = True
        End If
    End If
        
'        IsRegistered = True
    
    Load frmMain
    frmPictures.picQFG = frmSplash.imgLogo
    
    If IsThereRegistration = True Then
        Do
            If DateDiff("s", TheTime, Now) = 20 Then
                CanIGo = True
            End If
        Loop Until CanIGo = True
    End If
    
    frmSplash.Hide
    
    frmSplash.lblIrwinCo.Visible = True
    frmSplash.cmdOK.Visible = True
    frmSplash.cmdSysInfo.Visible = True
    frmSplash.lblWarning.Width = 369
    
    If IsThereRegistration = True Then
        If IsRegistered = True Then
            frmSplash.txtUserName.Locked = True
            frmSplash.txtUserNumber.Locked = True
            frmSplash.txtUserName.BackColor = &H8000000F
            frmSplash.txtUserNumber.BackColor = &H8000000F
        Else
            frmMain.chkPerfect.Enabled = False
            frmMain.mnuFilePerfect.Enabled = False
        End If
    Else
        frmMain.mnuHelpRegister.Visible = False
        frmMain.mnuLine3.Visible = False
        'frmMain.mnuFilePerfect.Visible = False
        'frmMain.mnuLine2.Visible = False
        'frmMain.chkPerfect.Visible = False
        
        frmMain.mnuFilePerfect.Visible = True
        frmMain.mnuLine2.Visible = True
        frmMain.chkPerfect.Visible = True
    End If
    frmMain.Show
End Sub
