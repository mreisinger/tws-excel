# Excel Add-In for Interactive Brokers TWS
This Excel Add-In provides an easy way to stream market data from Trader Workstation using the ActiveX API provided by Interactive Brokers.


At the moment these Worksheet functions are available:

sub_mktdata(id, symbol, secType, exchange, ccy)

  Subscribes to market data for the specified instrument.
  
    id: Unique identifier (integer) for the data stream
    symbol: IB Ticker for the instrument, e.g. "AAPL", "GOOGL"
    secType: Security type, e.g. "STK", "IND"
    exchange: Exchange, e.g. "SMART", "ISLAND", "IBIS"
    ccy: Currency, e.g. "USD", "EUR"
  
  
IBDP(id, datapoint)

  Returns the specified datapoint for the market data stream with identifier id.
  
    id: Unique identifier (integer) for the data stream
    datapoint: "bid", "ask", "last", "close", "bid_size", "ask_size"
