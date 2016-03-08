VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings 
   Caption         =   "Settings"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   OleObjectBlob   =   "Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Application.Calculation = xlCalculateManual
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value = autoConnect
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value = showError
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(6, 2).value = showStatus
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(7, 2).value = limitRefresh
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(8, 2).value = refreshRate.Text
    
    m_autoConnect = autoConnect
    m_showErrorMsgBox = showError
    m_showStatusBar = showStatus
    m_limitRefresh = limitRefresh
    
    If m_showStatusBar Then
        If TWS Is Nothing Then
            Application.StatusBar = "TWS not connected"
        Else
            If TWS.m_isConnected Then
                Application.StatusBar = "TWS connected"
            Else
                Application.StatusBar = "TWS not connected"
            End If
        End If
    Else
        Application.StatusBar = False
    End If
    
    Workbooks("TWS_API.xlam").Save
    Unload Settings
Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub CommandButton2_Click()
    Unload Settings
End Sub


Private Sub UserForm_Activate()
    Settings.Top = Application.Top + 300
    Settings.Left = Application.Left + 350
    
    If Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value <> "" Then
        Settings.autoConnect = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value
    Else
        Settings.autoConnect = True
    End If
    
    If Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value <> "" Then
        Settings.showError = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value
    Else
        Settings.showError = True
    End If
    
    If Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(6, 2).value <> "" Then
        Settings.showStatus = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(6, 2).value
    Else
        Settings.showStatus = True
    End If
    
    If Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(7, 2).value <> "" Then
        Settings.limitRefresh = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(7, 2).value
    Else
        Settings.limitRefresh = True
    End If
    
    refreshRate.Text = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(8, 2).value
    
End Sub
