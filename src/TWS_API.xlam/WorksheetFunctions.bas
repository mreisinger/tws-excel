Attribute VB_Name = "WorksheetFunctions"
Public Function sub_mktdata(id As Long, symbol As String, Optional secType As String = "STK", _
                            Optional exchange As String = "SMART", Optional curr As String = "USD", _
                            Optional expiry As String = "NOEXP", Optional c_p As String = "C", _
                            Optional strike As Double = 0, Optional multiplier As String = 100) As String

    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
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

            sub_mktdata = "ID " & id
        Else
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
    End If

End Function


Public Function IBDP(id As Long, datapoint As String) As Variant
Attribute IBDP.VB_Description = "Returns the specified data point. ""bid"", ""ask"", ""last"""

    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
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
        Else
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
    End If

End Function


Public Function req_ContractDetails(id As Long, secID As String, exchange As String) As String
    
    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
            
            'If arID(id, 1) = 0 Then
            
                Set TWS.m_contractInfo = TWS.m_TWSControl.createContract()
                
                With TWS.m_contractInfo
                    If Len(secID) = 12 Then
                        .secIdType = "ISIN"
                        .secID = secID
                    ElseIf Len(secID) = 6 Then
                        .secIdType = "ISIN"
                        .secID = assetCode("WKN", "ISIN", secID)
                    Else
                        req_ContractDetails = "Wrong Asset code"
                        Exit Function
                    End If
                End With
                
                Call TWS.m_TWSControl.reqContractDetailsEx(id, TWS.m_contractInfo)
            '    arID(id, 1) = 1
            'End If
            
            req_ContractDetails = "ID " & id
            
        Else
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
    End If

End Function


Public Function req_ContractDetailsWithTicker(id As Long, symbol As String, Optional secType As String = "STK", _
                            Optional exchange As String = "SMART", Optional curr As String = "USD", _
                            Optional expiry As String = "NOEXP", Optional c_p As String = "C", _
                            Optional strike As Double = 0, Optional multiplier As String = 100) As String
    
    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
            
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
            
            Call TWS.m_TWSControl.reqContractDetailsEx(id, TWS.m_contractInfo)
            
            req_ContractDetailsWithTicker = "ID " & id
            
        Else
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
    End If

End Function


Public Function displayContractDetails(id As Long, Optional transpose As Boolean = False) As Variant
    Dim details(25, 1) As Variant
    
    details(0, 0) = "ConID"
    details(1, 0) = "Symbol"
    details(2, 0) = "Security Type"
    details(3, 0) = "Expiry"
    details(4, 0) = "Strike"
    details(5, 0) = "Right"
    details(6, 0) = "Multiplier"
    details(7, 0) = "Exchange"
    details(8, 0) = "Primary Exchange"
    details(9, 0) = "Currency"
    details(10, 0) = "Local Symbol"
'   details(11, 0) = "Order Types"
    details(12, 0) = "Valid Exchanges"
    details(13, 0) = "Minimal Tick"
    details(14, 0) = "Market Name"
    details(15, 0) = "Trading Class"
    details(16, 0) = "Price Magnifier"
    details(17, 0) = "ev Rule"
    details(18, 0) = "ev Multiplierh"
    details(19, 0) = "Contract Month"
    details(20, 0) = "Industry"
    details(21, 0) = "Category"
    details(22, 0) = "Subcategory"
    details(23, 0) = "Time Zone"
    details(24, 0) = "Trading Hours"
    details(25, 0) = "Liquid Hours"
    
    With arConDetails(id)
        details(0, 1) = .m_conId
        details(1, 1) = .m_symbol
        details(2, 1) = .m_secType
        details(3, 1) = .m_lastTradeDateOrContractMonth
        details(4, 1) = .m_strike
        details(5, 1) = .m_right
        details(6, 1) = .m_multiplier
        details(7, 1) = .m_exchange
        details(8, 1) = .m_primaryExchange
        details(9, 1) = .m_currency
        details(10, 1) = .m_localSymbol
'        details(11, 1) = .m_orderTypes
        details(12, 1) = .m_validExchanges
        details(13, 1) = .m_minTick
        details(14, 1) = .m_marketName
        details(15, 1) = .m_tradingClass
        details(16, 1) = .m_priceMagnifier
        details(17, 1) = .m_evRule
        details(18, 1) = .m_evMultiplier
        details(19, 1) = .m_contractMonth
        details(20, 1) = .m_industry
        details(21, 1) = .m_category
        details(22, 1) = .m_subcategory
        details(23, 1) = .m_timeZoneId
        details(24, 1) = .m_tradingHours
        details(25, 1) = .m_liquidHours
    End With
    
    If transpose Then
        displayContractDetails = Application.transpose(details)
    Else
        displayContractDetails = details
    End If
    
End Function


Public Function IBDH(id As Long, symbol As String, Optional endDate As String = "", Optional secType As String = "STK", _
                            Optional exchange As String = "SMART", Optional curr As String = "USD", _
                            Optional expiry As String = "NOEXP", Optional c_p As String = "C", _
                            Optional strike As Double = 0, Optional multiplier As String = 100) As Variant
                            
    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
            Application.Volatile
            If refreshHistData Then
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
                
                If regexCompare(endDate, "^([1-9]|0[1-9]|[1-2][0-9]|3[0-1])\/([1-9]|0[1-9]|1[0-2])\/[1-2]\d{3}$") Then
                    endDateTime = Format(endDate, "YYYYMMDD") & " 23:59:59 GMT"
                ElseIf endDate = "" Then
                    endDateTime = Format(Now, "YYYYMMDD HH:mm:SS") + " GMT"
                Else
                    MsgBox "Wrong Date format"
                End If
                
                Duration = "1 D"
                barSize = "1 day"
                whatToShow = "TRADES"
                Dim useRTH As Long
                Dim formatDate As Long
                useRTH = 1
                formatDate = 1
        
                ' chart options
                Dim chartOptions As TWSLib.ITagValueList
                Set chartOptions = TWS.m_TWSControl.createTagValueList()
        
                ' call reqHistoricalDataEx method
                Call TWS.m_TWSControl.reqHistoricalDataEx(id, TWS.m_contractInfo, endDateTime, Duration, barSize, whatToShow, useRTH, formatDate, chartOptions)
            Else
                IBDH = arHistData(id).m_histClose
                refreshHistData = True
            End If
            
        Else
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
    End If

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
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
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
            MsgBox (str_not_connected)
        End If
    Else
        MsgBox (str_not_initialized)
    End If

End Function


Public Function assetCode(type1 As String, type2 As String, secID As String) As String

    If UCase(type1) = "WKN" Then
        If Len(secID) <> 6 Then
            assetCode = "Not a WKN (wrong length)"
            Exit Function
        End If
    End If
    
    If UCase(type1) = "ISIN" Then
        If Len(secID) <> 12 Then
            assetCode = "Not an ISIN (wrong length)"
            Exit Function
        End If
    End If

    If UCase(type1) = "WKN" And UCase(type2) = "ISIN" Then
        assetCode = wknToIsin(secID)
    End If
    
    If UCase(type1) = "ISIN" And UCase(type2) = "WKN" Then
        assetCode = isinToWkn(secID)
    End If

End Function
