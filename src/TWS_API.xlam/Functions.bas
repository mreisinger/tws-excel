Attribute VB_Name = "Functions"
Sub subscribe_mktdata(id As Long, symbol As String, secType As String, exchange As String, curr As String, _
                        expiry As String, c_p As String, strike As Double, multiplier As String)

    Set TWS.m_contractInfo = TWS.m_TWSControl.createContract()
    
    With TWS.m_contractInfo
        .symbol = UCase(symbol)
        .secType = UCase(secType)
        .exchange = UCase(exchange)
        '.primaryExchange = "IBIS"
        .currency = UCase(curr)
    End With
    
    If secType = "OPT" Or secType = "IOPT" Then
        With TWS.m_contractInfo
            .Right = UCase(c_p)
            .strike = strike
            .lastTradeDateOrContractMonth = expiry
            .multiplier = multiplier
        End With
    End If
    
    If secType = "FUT" Then
        With TWS.m_contractInfo
            .lastTradeDateOrContractMonth = expiry
        End With
    End If
    
    genericTickList = ""
    
    Dim mktDataOptions As TWSLib.ITagValueList
    Set mktDataOptions = TWS.m_TWSControl.createTagValueList()
    
    Call TWS.m_TWSControl.reqMktDataEx(id, TWS.m_contractInfo, genericTickList, 0, mktDataOptions)

End Sub

Public Sub UpdateArrayWithPrice(id As Long, tickType As Long, price As Double)
    
    With arMktData(id)
        Select Case tickType
            Case BID_PRICE
                .m_BidPrice = price
            Case ASK_PRICE
                .m_AskPrice = price
            Case LAST_PRICE
                .m_LastPrice = price
            Case CLOSE_PRICE
                .m_ClosePrice = price
        End Select
    End With
    
    ActiveSheet.Calculate
    
End Sub

Public Sub calc_sheet()

    ActiveSheet.Calculate
    allowRefresh = False

End Sub

Public Sub UpdateArrayWithSize(id As Long, tickType As Long, size As Long)

    With arMktData(id)
        Select Case tickType
            Case BID_SIZE
                .m_BidSize = size
            Case ASK_SIZE
                .m_AskSize = size
            Case LAST_SIZE
                .m_LastSize = size
        End Select
    End With
   
    ActiveSheet.Calculate
    
End Sub

Public Sub UpdateArrayWithString(id As Long, tickType As Long, value As String)

    With arMktData(id)
        Select Case tickType
            Case LAST_TIMESTAMP
                .m_LastTimeStamp = value
        End Select
    End With

    ActiveSheet.Calculate

End Sub
