VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrLog 
   Caption         =   "ErrorLog"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   OleObjectBlob   =   "ErrLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ErrLog.Hide
End Sub

Private Sub CommandButton2_Click()
    ErrLog.TextBox1.Text = ""
End Sub

Private Sub UserForm_Activate()
    ErrLog.Top = Application.Top + 300
    ErrLog.Left = Application.Left + 350
End Sub
