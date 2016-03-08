Attribute VB_Name = "Util"
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public TWS As cTWSControl
Public arID(500, 1) As Variant
Public arMktData(500) As mktDataRecord
Public arHistData(500) As histDataRecord
Public arConDetails(500) As conDetails
Public allowRefresh As Boolean
Public lastRefresh As Double
Public refreshHistData As Boolean

Public connectionHost As String
Public connectionPort As String
Public clientId As String

Public m_autoConnect As Boolean
Public m_showErrorMsgBox As Boolean
Public m_showStatusBar As Boolean
Public m_limitRefresh As Boolean
Public m_refreshRate As Integer

Public Const str_not_connected = "TWS is not connected"
Public Const str_not_initialized = "TWS Control is not initialized"

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type mktDataRecord
    m_secType As String
    m_BidPrice As Double
    m_BidSize As Double
    m_AskPrice As Double
    m_AskSize As Double
    m_LastPrice As Double
    m_LastSize As Double
    m_ClosePrice As Double
    m_LastTimeStamp As String
End Type

Public Type histDataRecord
    m_histDate As String
    m_histOpen As Double
    m_histHigh As Double
    m_histLow As Double
    m_histClose As Double
    m_histVolume As Long
    m_barCount As Long
    m_WAP As Double
    m_hasGaps As Long
End Type

Public Type conDetails
    m_conId As Long
    m_symbol As String
    m_secType As String
    m_lastTradeDateOrContractMonth As String
    m_strike As Double
    m_right As String
    m_multiplier As String
    m_exchange As String
    m_primaryExchange As String
    m_currency As String
    m_localSymbol As String
    m_orderTypes As String
    m_validExchanges As String
    m_minTick As Double
    m_marketName As String
    m_tradingClass As String
    m_priceMagnifier As Long
    m_evRule As String
    m_evMultiplier As Double
    m_contractMonth As String
    m_industry As String
    m_category As String
    m_subcategory As String
    m_timeZoneId As String
    m_tradingHours As String
    m_liquidHours As String
End Type

Public Enum tickType
    BID_SIZE
    BID_PRICE
    ASK_PRICE
    ASK_SIZE
    LAST_PRICE
    LAST_SIZE
    HIGH_TICK
    LOW_TICK
    VOLUME_TICK
    CLOSE_PRICE
    BID_OPTION_COMPUTATION
    ASK_OPTION_COMPUTATION
    LAST_OPTION_COMPUTATION
    MODEL_OPTION
    OPEN_TICK
    LOW_13_WEEK
    HIGH_13_WEEK
    LOW_26_WEEK
    HIGH_26_WEEK
    LOW_52_WEEK
    HIGH_52_WEEK
    AVG_VOLUME
    OPEN_INTEREST
    OPTION_HISTORICAL_VOL
    OPTION_IMPLIED_VOL
    OPTION_BID_EXCH
    OPTION_ASK_EXCH
    OPTION_CALL_OPEN_INTEREST
    OPTION_PUT_OPEN_INTEREST
    OPTION_CALL_VOLUME
    OPTION_PUT_VOLUME
    INDEX_FUTURE_PREMIUM
    BID_EXCH
    ASK_EXCH
    AUCTION_VOLUME
    AUCTION_PRICE
    AUCTION_IMBALANCE
    MARK_PRICE
    BID_EFP_COMPUTATION
    ASK_EFP_COMPUTATION
    LAST_EFP_COMPUTATION
    OPEN_EFP_COMPUTATION
    HIGH_EFP_COMPUTATION
    LOW_EFP_COMPUTATION
    CLOSE_EFP_COMPUTATION
    LAST_TIMESTAMP
    SHORTABLE
    FUNDAMENTAL_RATIOS
    RT_VOLUME
    HALTED
    BID_YIELD
    ASK_YIELD
    LAST_YIELD
    CUST_OPTION_COMPUTATION
    TRADE_COUNT
    TRADE_RATE
    VOLUME_RATE
    LAST_RTH_TRADE
End Enum
