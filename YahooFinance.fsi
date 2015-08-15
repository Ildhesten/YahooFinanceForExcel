namespace YahooFinanceForExcel

  module YahooFinance =    
    open System

    open ExcelDna.Integration

    type t = Dummy

    [<ExcelFunction(Description="Gets quote data from Yahoo Finance")>]
    val YahooGetQuoteDetails : symbol : string -> obj[,]

    [<ExcelFunction(Description="Gets historical data from Yahoo Finance")>]    
    val YahooGetHistoricalPrices : symbol:string -> granularity:string -> startDate:DateTime -> endDate:DateTime -> obj[,]

    [<ExcelFunction(Description="Gets a sector overview from Yahoo Finance")>]   
    val YahooGetSectorOverview : unit -> obj[,]