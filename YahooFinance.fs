namespace YahooFinanceForExcel

module YahooFinance =
  open System.IO
  open System.Net
  open System.Web
  open System

  open ExcelDna.Integration

  type t = Dummy

  let findLongestRow (rows : string [] []) =
    Array.fold (fun a (x : string[]) -> max a x.Length) 0 rows

  let readFromArray (array : string[][]) i j =
    if array.Length > i then
      if array.[i].Length > j then
        array.[i].[j] :> obj
      else
        null
    else
      null

  let convertToArray2d (getInput : unit -> string[][]) =
    try
      let input = getInput ()
      in
        Array2D.init input.Length (findLongestRow input) (readFromArray input)
    with
      exn -> Array2D.init 1 1 (fun _ _ -> exn.Message :> obj)

  let GetCsvDataExn (url : string) init =
    let request = WebRequest.Create(url)
    use response = request.GetResponse()
    use reader = new StreamReader(response.GetResponseStream())
    let mutable lines = init
    in
      begin
        while not reader.EndOfStream do
          lines <- reader.ReadLine() :: lines
        List.rev lines
        |> List.map (fun (s : string) -> s.Split([|','|]))
        |> List.toArray
      end


  let YahooGetHistoricalPricesExn (symbol : string) (granularity : string) (startDate : DateTime) (endDate : DateTime) =
    let encodedUrl = HttpUtility.UrlEncode(symbol)
    let encodedGranularity = HttpUtility.UrlEncode(granularity)
    let requestString = sprintf "http://real-chart.finance.yahoo.com/table.csv?s=%s&g=%s&a=%d&b=%d&c=%d&d=%d&e=%d&f=%d" encodedUrl granularity (startDate.Month - 1) startDate.Day startDate.Year (endDate.Month - 1) endDate.Day endDate.Year
    in
      GetCsvDataExn requestString []

  let YahooGetSectorOverviewExn () =
    let requestString = "http://biz.yahoo.com/p/csv/s_conameu.csv"
    in
      GetCsvDataExn requestString []

  [<ExcelFunction(Description="Gets historical data from YahooFinance.")>]
  let YahooGetHistoricalPrices 
    ([<ExcelArgument(Name="Symbol", Description="The symbol to look up on Yahoo Finance.")>] symbol : string) 
    ([<ExcelArgument(Name="Symbol", Description="d for daily, w for weekly, m for monthly.")>] granularity : string) 
    ([<ExcelArgument(Name="StartDate", Description="The start date of the range.")>] startDate : DateTime) 
    ([<ExcelArgument(Name="EndDate", Description="The end date of the range")>] endDate : DateTime) =
      convertToArray2d (fun () -> YahooGetHistoricalPricesExn symbol granularity startDate endDate)            

  [<ExcelFunction(Description="Gets a sector overview from YahooFinance.")>]
  let YahooGetSectorOverview () =
    convertToArray2d YahooGetSectorOverviewExn

  let YahooGetQuoteDetailsExn (symbol : string) =    
    let escapedSymbol = HttpUtility.UrlEncode(symbol)
    let requestString = sprintf "http://download.finance.yahoo.com/d/quotes.csv?s=%s&f=nsl1oph0g0" escapedSymbol    
    let lines = ["Name,Symbol,Last,Open,Prev. Close,High,Low"]
    in
      GetCsvDataExn requestString lines

  [<ExcelFunction(Description="Gets a sector overview from YahooFinance")>]
  let YahooGetQuoteDetails symbol =
    convertToArray2d (fun () -> YahooGetQuoteDetailsExn symbol)