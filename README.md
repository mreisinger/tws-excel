# Excel Add-In for Interactive Brokers TWS
This Excel Add-In provides an easy way to stream market data from Trader Workstation using the ActiveX API provided by Interactive Brokers.

## How to install:

- Enable ActiveX Clients in TWS API settings
- Import Add-In in Excel
- Change connection settings (in ribbon "TWS API")


## Available worksheet functions:

#### =sub_mktdata(id, localSymbol, secType, exchange)

  Subscribes to market data for the specified instrument
  
    id:           Unique identifier (integer) for the data stream
    localSymbol:  Local symbol for the instrument, e.g.
                  STK: "AAPL", "BMW"
                  FUT: "6EH6"
                  FOP: "6E2H6 C1085"
                  OPT: "C ODAX MAR 16  9550"
                  WAR: "GL1806" (WKN)
    secType:      Security type, e.g. "STK", "OPT", "FUT"
    exchange:     Exchange, e.g. "SMART", "ISLAND", "IBIS"

#### =sub_mktdataWithTicker(id, symbol, secType, exchange, ccy, optional expiry, optional right, optional strike, optional multiplier)

  Subscribes to market data for the specified instrument.
  
##### Stocks
  
    id:         Unique identifier (integer) for the data stream
    symbol:     IB Ticker for the instrument, e.g. "AAPL", "GOOGL"
    secType:    Security type, e.g. "STK", "OPT", "FUT"
    exchange:   Exchange, e.g. "SMART", "ISLAND", "IBIS"
    ccy:        Currency, e.g. "USD", "EUR"
  
##### Futures:
  
    In addition to parameters listed under stocks:
    expiry:     Expiry month, "YYYYMM"

##### Options
  
    In addition to parameters listed under stocks:
    expiry:     Expiry date, "YYYYMMDD"
    right:      Call or Put: "C", "P"
    strike:     Strike price
    multiplier: Multiplier, e.g. 100
  
#### =IBDP(id, datapoint)

  Returns the specified datapoint for the market data stream with identifier id.
  
    id:         Unique identifier (integer) for the data stream
    datapoint:  "bid", "ask", "last", "close", "bid_size", "ask_size"

#### =req_ContractDetails(id, localSymbol, secType, exchange)

  Downloads contract details for the specified instrument.
  
    id:           Unique identifier (integer) for the data stream
    localSymbol:  ISIN, CUSIP, or local symbol (see submktdata for details)
    secType:      Security type, e.g. "STK", "OPT", "FUT"
    exchange:     Exchange, e.g. "SMART", "ISLAND", "IBIS"
  
#### =req_ContractDetailsWithTicker(id, symbol, secType, exchange, ccy, optional expiry, optional right, optional strike, optional multiplier)

  Downloads contract details for the specified instrument.
  
##### Stocks
  
    id:         Unique identifier (integer) for the data stream
    symbol:     IB Ticker for the instrument, e.g. "AAPL", "GOOGL"
    secType:    Security type, e.g. "STK", "OPT", "FUT"
    exchange:   Exchange, e.g. "SMART", "ISLAND", "IBIS"
    ccy:        Currency, e.g. "USD", "EUR"
  
##### Futures:
  
    In addition to parameters listed under stocks:
    expiry:     Expiry month, "YYYYMM"

##### Options
  
    In addition to parameters listed under stocks:
    expiry:     Expiry date, "YYYYMMDD"
    right:      Call or Put: "C", "P"
    strike:     Strike price
    multiplier: Multiplier, e.g. "100"

#### =displayContractDetails(id, transpose)

  Returns an array with contract details
  
    id:         Unique identifier (integer) for the data stream
    transpose:  "TRUE" or "FALSE"

#### =assetCode(type1, type2, secID)

  Converts asset codes
  
    type1:      "ISIN" or "WKN"
    type2:      "ISIN" or "WKN"
    secID:      ISIN or WKN, e.g. "US0378331005"

#### =cancel_mktdata(id)

  Cancels the market data stream with identifier id.
  
    id:         Unique identifier (integer) for the data stream

## Currently working on

  - More datapoints, generic ticks
  - Option chain
  - Option strategies
  - Account and Portfolio details
  - Error handling
