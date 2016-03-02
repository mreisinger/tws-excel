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
            MsgBox ("TWS not connected")
        End If
    Else
        MsgBox ("TWSControl not initialized")
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
            MsgBox ("TWS not connected")
        End If
    Else
        MsgBox ("TWSControl not initialized")
    End If

End Function


Public Function req_ContractDetails(id As Long) As Variant
    req_ContractDetails = details
    Application.Volatile
    If Not (TWS Is Nothing) Then
        If TWS.m_isConnected Then
            
            If arID(id, 1) = 0 Then
            
                Set TWS.m_contractInfo = TWS.m_TWSControl.createContract()
                
                With TWS.m_contractInfo
                    .secIdType = "ISIN"
                    .secID = "US0378331005" 'AssetCode("WKN", "ISIN", "CC4UAE")
    '                .symbol = UCase(Cells(id, Columns(COLUMN_SYMBOL).Column).value)
    '                .secType = UCase(Cells(id, Columns(COLUMN_SECTYPE).Column).value)
    '                .lastTradeDateOrContractMonth = Cells(id, Columns(COLUMN_LASTTRADEDATE).Column).value
    '                .strike = Cells(id, Columns(COLUMN_STRIKE).Column).value
    '                .Right = UCase(Cells(id, Columns(COLUMN_RIGHT).Column).value)
    '                .multiplier = UCase(Cells(id, Columns(COLUMN_MULTIPLIER).Column).value)
                    .exchange = "SMART"
    '                .primaryExchange = "ISLAND"
    '                .currency = UCase(Cells(id, Columns(COLUMN_CURRENCY).Column).value)
    '                .localSymbol = UCase(Cells(id, Columns(COLUMN_LOCALSYMBOL).Column).value)
    '                .includeExpired = Cells(id, Columns(COLUMN_INCLUDEEXPIRED).Column).value
                End With
                
                Call TWS.m_TWSControl.reqContractDetailsEx(id, TWS.m_contractInfo)
                arID(id, 1) = 1
            End If
            
            'Debug.Print arConDetails(id).m_symbol
'            With arConDetails(id)
'                details(0, 1) = .m_conId
'                details(1, 1) = .m_symbol
'                details(2, 1) = .m_secType
'                details(3, 1) = .m_lastTradeDateOrContractMonth
'                details(4, 1) = .m_strike
'                details(5, 1) = .m_right
'                details(6, 1) = .m_multiplier
'                details(7, 1) = .m_exchange
'                details(8, 1) = .m_primaryExchange
'                details(9, 1) = .m_currency
'                details(10, 1) = .m_localSymbol
'                details(11, 1) = .m_orderTypes
'                details(12, 1) = .m_validExchanges
'                details(13, 1) = .m_minTick
'                details(14, 1) = .m_marketName
'                'details(0,1) = .m_tradingClass
'                details(15, 1) = .m_priceMagnifier
'                details(16, 1) = .m_evRule
'                details(17, 1) = .m_evMultiplier
'                details(19, 1) = .m_contractMonth
'                details(20, 1) = .m_industry
'                details(21, 1) = .m_category
'                details(22, 1) = .m_subcategory
'                details(23, 1) = .m_timeZoneId
'                details(24, 1) = .m_tradingHours
'                details(25, 1) = .m_liquidHours
'            End With
            
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
'            details(11, 0) = "Order Types"
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
            
            req_ContractDetails = details
            
        Else
            MsgBox ("TWS not connected")
        End If
    Else
        MsgBox ("TWSControl not initialized")
    End If

End Function

Sub wetr()
For i = 0 To 17
    Debug.Print details(i, 1)
Next i

End Sub


Public Function test() As Variant
    Application.Volatile
    test = details
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


Public Function AssetCode(type1 As String, type2 As String, secID As String) As String

    If UCase(type1) = "WKN" Then
        If Len(secID) <> 6 Then
            AssetCode = "Not a WKN (wrong length)"
            Exit Function
        End If
    End If
    
    If UCase(type1) = "ISIN" Then
        If Len(secID) <> 12 Then
            AssetCode = "Not an ISIN (wrong length)"
            Exit Function
        End If
    End If

    If UCase(type1) = "WKN" And UCase(type2) = "ISIN" Then
        AssetCode = wknToIsin(secID)
    End If
    
    If UCase(type1) = "ISIN" And UCase(type2) = "WKN" Then
        AssetCode = isinToWkn(secID)
    End If

End Function
