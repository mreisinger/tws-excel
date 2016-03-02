Attribute VB_Name = "IncomingTicks"
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


Public Sub UpdateContractDetails(ByVal reqId As Long, ByVal contractDetails As TWSLib.IContractDetails)
    
    With arConDetails(reqId)
    'Debug.Print reqId
    'Debug.Print contractDetails.Summary.symbol
    'Debug.Print contractDetails.Summary.conId
    'Debug.Print contractDetails.Summary.secType
        .m_conId = contractDetails.Summary.conId
        .m_symbol = contractDetails.Summary.symbol
'        .m_secType = contractDetails.Summary.secType
'        .m_lastTradeDateOrContractMonth = contractDetails.Summary.lastTradeDateOrContractMonth
'        .m_strike = contractDetails.Summary.strike
'        .m_right = contractDetails.Summary.Right
'        .m_multiplier = contractDetails.Summary.multiplier
'        .m_exchange = contractDetails.Summary.exchange
'        .m_primaryExchange = contractDetails.Summary.primaryExchange
'        .m_currency = contractDetails.Summary.currency
'        .m_localSymbol = contractDetails.Summary.localSymbol
'        '.m_orderTypes = contractDetails.orderTypes
'        .m_validExchanges = contractDetails.validExchanges
'        .m_minTick = contractDetails.minTick
'        .m_marketName = contractDetails.marketName
'        .m_tradingClass = contractDetails.Summary.tradingClass
'        .m_priceMagnifier = contractDetails.priceMagnifier
'        .m_evRule = contractDetails.evRule
'        .m_evMultiplier = contractDetails.evMultiplier
'        .m_contractMonth = contractDetails.evMultiplier
'        .m_industry = contractDetails.industry
'        .m_category = contractDetails.Category
'        .m_subcategory = contractDetails.subcategory
'        .m_timeZoneId = contractDetails.timeZoneId
'        .m_tradingHours = contractDetails.tradingHours
'        .m_liquidHours = contractDetails.liquidHours
        
        
        details(0, 1) = .m_conId
        details(1, 1) = .m_symbol
'        details(2, 1) = .m_secType
'        details(3, 1) = .m_lastTradeDateOrContractMonth
'        details(4, 1) = .m_strike
'        details(5, 1) = .m_right
'        details(6, 1) = .m_multiplier
'        details(7, 1) = .m_exchange
'        details(8, 1) = .m_primaryExchange
'        details(9, 1) = .m_currency
'        details(10, 1) = .m_localSymbol
''        details(11, 1) = .m_orderTypes
'        details(12, 1) = .m_validExchanges
'        details(13, 1) = .m_minTick
'        details(14, 1) = .m_marketName
'        details(15, 1) = .m_tradingClass
'        details(16, 1) = .m_priceMagnifier
'        details(17, 1) = .m_evRule
'        details(18, 1) = .m_evMultiplier
'        details(19, 1) = .m_contractMonth
'        details(20, 1) = .m_industry
'        details(21, 1) = .m_category
'        details(22, 1) = .m_subcategory
'        details(23, 1) = .m_timeZoneId
'        details(24, 1) = .m_tradingHours
'        details(25, 1) = .m_liquidHours
    End With

End Sub


Public Sub calc_sheet()

    ActiveSheet.Calculate
    allowRefresh = False

End Sub
