Attribute VB_Name = "Ribbon"
Public objRibbon As IRibbonUI

Public Sub onload(Ribbon As IRibbonUI)
    Set objRibbon = Ribbon
End Sub


Public Function TWS_Connect(control As IRibbonControl)
Attribute TWS_Connect.VB_Description = "Connects to TWS. Port is hardcoded."

    getConnectionDetails
    
    If connectionPort = "" Or clientId = "" Then
        MsgBox "Please check connection details"
        port.Show
        Exit Function
    End If
    
    If TWS Is Nothing Then
        Set TWS = New cTWSControl
    End If
    
    If TWS Is Nothing Then
        MsgBox ("TWSControl not initialized")
    Else
        If Not TWS.m_isConnected Then
            Call TWS.m_TWSControl.Connect(connectionHost, connectionPort, clientId, False)
            TWS.m_isConnected = True
        Else
            MsgBox ("Already connected")
        End If
    End If
    Application.StatusBar = "TWS connected"
    
End Function


Public Function TWS_Disconnect(control As IRibbonControl)

    If Not (TWS Is Nothing) And TWS.m_isConnected Then
        Call TWS.m_TWSControl.Disconnect
        TWS.m_isConnected = False
    Else
        MsgBox ("Not connected")
    End If
    Application.StatusBar = "TWS not connected"
    
End Function


Public Sub getConnectionDetails()

    connectionHost = GetSetting("Microsoft Excel", "TWS API", "Host", "")
    connectionPort = GetSetting("Microsoft Excel", "TWS API", "Port", "")
    clientId = GetSetting("Microsoft Excel", "TWS API", "ClientID", "")

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
