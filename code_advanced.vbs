Sub StockMarketAdvanced()

  For Each ws In Worksheets

  ' Set variables
  Dim Range_Percent As Range
  Dim Range_Volume As Range
  Dim Percent_Max As Double
  Dim Percent_Min As Double
  Dim Volume_Max As Double
  Dim PerMaxTicker As String
  Dim PerMinTicker As String
  Dim VolMaxTicker As String
  
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
  'Set ranges
  Set Range_Percent = ws.Range("K2:K" & LastRow)
  Set Range_Volume = ws.Range("L2:L" & LastRow)
  
  'Determine min or max
  RngPerMin = Application.WorksheetFunction.Min(Range_Percent)
  RngPerMax = Application.WorksheetFunction.Max(Range_Percent)
  RngVolMax = Application.WorksheetFunction.Max(Range_Volume)
  
  PerMaxTicker = Application.Index(ws.Range("I:I"), Application.Match(RngPerMax, ws.Range("K:K"), 0))
  PerMinTicker = Application.Index(ws.Range("I:I"), Application.Match(RngPerMin, ws.Range("K:K"), 0))
  VolMaxTicker = Application.Index(ws.Range("I:I"), Application.Match(RngVolMax, ws.Range("L:L"), 0))


        'Put headers and titles
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        'Print max/min values
        ws.Range("Q2") = RngPerMax
        ws.Range("Q3") = RngPerMin
        ws.Range("Q4") = RngVolMax
        
        'Print belonging Tickers
        ws.Range("P2") = PerMaxTicker
        ws.Range("P3") = PerMinTicker
        ws.Range("P4") = VolMaxTicker
        


        ws.Range("Q2:Q3").Style = "Percent"
        ws.Columns("O:Q").AutoFit
        
Next ws

End Sub

