VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcelEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Dim ctl As IRibbonControl
    
    m_autoConnect = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value
    
    If Wb.Name = "TWS_API.xlam" And m_autoConnect Then
        TWS_Connect ctl
    End If
End Sub
