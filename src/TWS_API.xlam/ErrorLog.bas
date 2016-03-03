Attribute VB_Name = "ErrorLog"
Public Sub LogMessage(ByVal id As Long, ByVal errorCode As Long, ByVal errorMsg As String)

    ErrLog.TextBox1.Text = ErrLog.TextBox1.Text & vbNewLine & Now & "  --  Error " & errorCode & " : " & errorMsg
    
    If Settings.showError Then
        If errorCode <> 2104 And errorCode <> 2106 Then
            MsgBox ("ID: " & id & "  " & "Code: " & errorCode & ": " & errorMsg)
        End If
    End If
    
End Sub
