Attribute VB_Name = "Util"
Public TWS As cTWSControl
Public arMktData(500) As mktDataRecord
Public allowRefresh As Boolean
Public starttime As Date

Public connectionHost As String
Public connectionPort As String
Public clientId As String

Public Type mktDataRecord
    m_secType As String
    
    ' tick price
    m_BidPrice As Double
    m_BidSize As Double
    m_AskPrice As Double
    m_AskSize As Double
    m_LastPrice As Double
    m_LastSize As Double
    m_ClosePrice As Double
    m_LastTimeStamp As String
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
