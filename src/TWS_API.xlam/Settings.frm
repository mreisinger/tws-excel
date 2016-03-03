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
    SaveSetting "Microsoft Excel", "TWS API", "ShowError", showError
    Settings.Hide
End Sub

Private Sub CommandButton2_Click()
    Settings.Hide
End Sub

Private Sub UserForm_Activate()
    Settings.Top = Application.Top + 300
    Settings.Left = Application.Left + 350
    
    Settings.showError = GetSetting("Microsoft Excel", "TWS API", "ShowError", "")
End Sub
