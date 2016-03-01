Attribute VB_Name = "WorksheetFunctions"
Public Function sub_mktdata(id As Long, symbol As String, Optional secType As String = "STK", _
                            Optional exchange As String = "SMART", Optional curr As String = "USD", _
                            Optional expiry As String = "NOEXP", Optional c_p As String = "C", _
                            Optional strike As Double = 0, Optional multiplier As String = 100) As String

    Call subscribe_mktdata(id, symbol, secType, exchange, curr, expiry, c_p, strike, multiplier)

    sub_mktdata = "ID " & id

End Function


Public Function IBDP(id As Long, datapoint As String) As Variant
Attribute IBDP.VB_Description = "Returns the specified data point. ""bid"", ""ask"", ""last"""

    Application.Volatile
    With arMktData(id)
        Select Case datapoint
            Case "bid"
                IBDP = .m_BidPrice
            Case "bid_size"
                IBDP = .m_BidSize
            Case "ask"
                IBDP = .m_AskPrice
            Case "ask_size"
                IBDP = .m_AskSize
            Case "last"
                IBDP = .m_LastPrice
            Case "last_size"
                IBDP = .m_LastSize
            Case "close"
                IBDP = .m_ClosePrice
        End Select
    End With

End Function


Public Function cancel_mktdata(id As Integer) As String

    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
            Call TWS.m_TWSControl.cancelMktData(id)
            
            ' clear array
            With arMktData(id)
                .m_BidPrice = 0
                .m_AskPrice = 0
                .m_LastPrice = 0
                .m_ClosePrice = 0
                .m_BidSize = 0
                .m_AskSize = 0
                .m_LastSize = 0
                .m_LastTimeStamp = ""
            End With
            
            cancel_mktdata = "ID " & id & " canceled"
        Else
            MsgBox ("TWS not connected")
        End If
    Else
        MsgBox ("TWSControl not initialized")
    End If

End Function

Public Function cancel_mktdata_all() As String

    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
            For id = 1 To 200
                Call TWS.m_TWSControl.cancelMktData(id)
                
                ' clear array
                With arMktData(id)
                    .m_BidPrice = 0
                    .m_AskPrice = 0
                    .m_LastPrice = 0
                    .m_ClosePrice = 0
                    .m_BidSize = 0
                    .m_AskSize = 0
                    .m_LastSize = 0
                    .m_LastTimeStamp = ""
                End With
            Next id
            
            cancel_all_mktdata = "All canceled"
        Else
            MsgBox ("TWS not connected")
        End If
    Else
        MsgBox ("TWSControl not initialized")
    End If

End Function
