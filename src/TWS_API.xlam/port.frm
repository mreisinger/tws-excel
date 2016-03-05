VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Port 
   Caption         =   "Connection Details"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   OleObjectBlob   =   "port.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "port"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(1, 2).value = TextBox1.Text
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(2, 2).value = TextBox2.Text
    Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(3, 2).value = TextBox3.Text
    
    Workbooks("TWS_API.xlam").Save
    Unload port
End Sub

Private Sub CommandButton2_Click()
    Unload port
End Sub

Private Sub UserForm_Activate()
    port.Top = Application.Top + 300
    port.Left = Application.Left + 500
    
    TextBox1.Text = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(1, 2).value
    TextBox2.Text = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(2, 2).value
    TextBox3.Text = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(3, 2).value
End Sub

