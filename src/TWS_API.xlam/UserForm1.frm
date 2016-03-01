VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Connection Details"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    SaveSetting "Microsoft Excel", "TWS API", "Host", TextBox1.Text
    SaveSetting "Microsoft Excel", "TWS API", "Port", TextBox2.Text
    SaveSetting "Microsoft Excel", "TWS API", "ClientID", TextBox3.Text
    Unload UserForm1
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm1
End Sub

Private Sub UserForm_Activate()
    UserForm1.Top = Application.Top + 300
    UserForm1.Left = Application.Left + 500
    
    TextBox1.Text = GetSetting("Microsoft Excel", "TWS API", "Host", "")
    TextBox2.Text = GetSetting("Microsoft Excel", "TWS API", "Port", "")
    TextBox3.Text = GetSetting("Microsoft Excel", "TWS API", "ClientID", "")
End Sub

