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
    
    If m_limitRefresh Then
        calc_sheet
    Else
        ActiveSheet.Calculate
    End If
    
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
   
    'ActiveSheet.Calculate
    'calc_sheet
    
End Sub


Public Sub UpdateArrayWithString(id As Long, tickType As Long, value As String)

    With arMktData(id)
        Select Case tickType
            Case LAST_TIMESTAMP
                .m_LastTimeStamp = value
        End Select
    End With

    'ActiveSheet.Calculate

End Sub


Public Sub UpdateContractDetails(ByVal reqId As Long, ByVal contractDetails As TWSLib.IContractDetails)
    
    With arConDetails(reqId)
        .m_conId = contractDetails.Summary.conId
        .m_symbol = contractDetails.Summary.symbol
        .m_secType = contractDetails.Summary.secType
        .m_lastTradeDateOrContractMonth = contractDetails.Summary.lastTradeDateOrContractMonth
        .m_strike = contractDetails.Summary.strike
        .m_right = contractDetails.Summary.Right
        .m_multiplier = contractDetails.Summary.multiplier
        .m_exchange = contractDetails.Summary.exchange
        .m_primaryExchange = contractDetails.Summary.primaryExchange
        .m_currency = contractDetails.Summary.currency
        .m_localSymbol = contractDetails.Summary.localSymbol
        '.m_orderTypes = contractDetails.orderTypes
        .m_validExchanges = contractDetails.validExchanges
        .m_minTick = contractDetails.minTick
        .m_marketName = contractDetails.marketName
        .m_tradingClass = contractDetails.Summary.tradingClass
        .m_priceMagnifier = contractDetails.priceMagnifier
        .m_evRule = contractDetails.evRule
        .m_evMultiplier = contractDetails.evMultiplier
        .m_contractMonth = contractDetails.contractMonth
        .m_industry = contractDetails.industry
        .m_category = contractDetails.Category
        .m_subcategory = contractDetails.subcategory
        .m_timeZoneId = contractDetails.timeZoneId
        .m_tradingHours = contractDetails.tradingHours
        .m_liquidHours = contractDetails.liquidHours
    End With

End Sub


Public Sub UpdateHistoricalData(id As Long, histDate As String, histOpen As Double, histHigh As Double, _
                                histLow As Double, histClose As Double, histVolume As Long, barCount As Long, _
                                WAP As Double, hasGaps As Long)
    
    With arHistData(id)
        .m_histDate = histDate
        .m_histOpen = histOpen
        .m_histHigh = histHigh
        .m_histLow = histLow
        .m_histClose = histClose
        .m_histVolume = histVolume
        .m_barCount = barCount
        .m_WAP = WAP
        .m_hasGaps = hasGaps
    End With
    
    refreshHistData = False
    ActiveSheet.Calculate
    
End Sub


Public Sub calc_sheet()

    Dim sysTime As SYSTEMTIME
    GetSystemTime sysTime
    
    
    If (Format(sysTime.wMinute, "00") & Format(sysTime.wSecond, "00") & Format(sysTime.wMilliseconds, "000")) > (lastRefresh + m_refreshRate * 1000) Then
        ActiveSheet.Calculate
        lastRefresh = Format(sysTime.wMinute, "00") & Format(sysTime.wSecond, "00") & Format(sysTime.wMilliseconds, "000")
        Debug.Print lastRefresh
    End If
    'allowRefresh = False
    
End Sub
