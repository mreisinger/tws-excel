Attribute VB_Name = "Ribbon"
Public objRibbon As IRibbonUI

Public Sub onload(Ribbon As IRibbonUI)
    Set objRibbon = Ribbon
End Sub


Public Function TWS_Connect(control As IRibbonControl)
Attribute TWS_Connect.VB_Description = "Connects to TWS. Port is hardcoded."

    getConnectionDetails
    m_showErrorMsgBox = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(4, 2).value
    m_showStatusBar = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(5, 2).value
    m_limitRefresh = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(7, 2).value
    m_refreshRate = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(8, 2).value
    
    If connectionPort = "" Or clientId = "" Then
        MsgBox "Please check connection details"
        port.Show
        Exit Function
    End If
    
    If TWS Is Nothing Then
        Set TWS = New cTWSControl
    End If
    
    If TWS Is Nothing Then
        MsgBox (str_not_initialized)
    Else
        If Not TWS.m_isConnected Then
            Call TWS.m_TWSControl.Connect(connectionHost, connectionPort, clientId, False)
            TWS.m_isConnected = True
        Else
            MsgBox ("Already connected")
        End If
    End If
    
    If m_showStatusBar Then
        Application.StatusBar = "TWS connected"
    End If
    
End Function


Public Function TWS_Disconnect(control As IRibbonControl)

    If Not (TWS Is Nothing) And TWS.m_isConnected Then
        Call TWS.m_TWSControl.Disconnect
        TWS.m_isConnected = False
    Else
        MsgBox (str_not_connected)
    End If
    
    If m_showStatusBar Then
        Application.StatusBar = "TWS not connected"
    End If
    
End Function


Public Sub getConnectionDetails()

    connectionHost = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(1, 2).value
    connectionPort = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(2, 2).value
    clientId = Workbooks("TWS_API.xlam").Sheets("Sheet1").Cells(3, 2).value

End Sub


Public Sub change_port(control As IRibbonControl)

    port.Show

End Sub


Public Sub show_log(control As IRibbonControl)

    ErrLog.Show
    ErrLog.TextBox1.SetFocus

End Sub


Public Sub show_settings(control As IRibbonControl)

    Settings.Show

End Sub
