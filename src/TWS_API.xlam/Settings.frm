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
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value = showError
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value = showStatus
    
    m_showErrorMsgBox = showError
    m_showStatusBar = showStatus
    
    If m_showStatusBar Then
        If TWS.m_isConnected Then
            Application.StatusBar = "TWS connected"
        Else
            Application.StatusBar = "TWS not connected"
        End If
    Else
        Application.StatusBar = False
    End If
    
    Workbooks("TWS_API.xlam").Save
    Unload Settings
End Sub


Private Sub CommandButton2_Click()
    Unload Settings
End Sub


Private Sub UserForm_Activate()
    Settings.Top = Application.Top + 300
    Settings.Left = Application.Left + 350
    
    If Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value <> "" Then
        Settings.showError = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value
    Else
        Settings.showError = True
    End If
    
    If Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value <> "" Then
        Settings.showStatus = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value
    Else
        Settings.showStatus = True
    End If
End Sub
