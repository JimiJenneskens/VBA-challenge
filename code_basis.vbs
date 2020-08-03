Sub StockMarket()

  For Each ws In Worksheets
  
    ' Set an initial variable for holding the ticker, opening value and closing value
  Dim Ticker As String
  Dim Ticket_Total As Long
  Dim Ticker_Opening As Double
  Dim Ticker_Closing As Double
  Dim Yearly_Change As Double
  Dim Percentage_Change As Double
  
  ' Set an initial variable for holding the total per ticker
  Dim Ticker_Total As Double
  Ticker_Total = 0
  
    ' Determine the Last Row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all tickers
  For i = 2 To LastRow
  
    ' Check if previous ticker is different than current to determine Ticker Opening
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Opening
      Ticker_Opening = ws.Cells(i, 3).Value

    Else
    End If

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value
      
      ' Add to the Ticket Total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
      
      ' Set Ticker Closing
      Ticker_Closing = ws.Cells(i, 6).Value
      
      ' Set Yearly Change
      Yearly_Change = Ticker_Closing - Ticker_Opening
      
      ' Set Percentage Change
      Percentage_Change = (Ticker_Closing - Ticker_Opening) / Ticker_Opening

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      
      ' Print the Ticker Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
      
      ' Print the Yearly Change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print the Yearly Change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percentage_Change

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the Ticker Total
      Ticker_Total = 0
      
        ' Add Titles to summary table
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
    
    ' Determine if Yearly Change is positive or negative
    If Yearly_Change > 0 Then
        ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
    ElseIf Yearly_Change < 0 Then
        ws.Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
        
    End If

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If
       

  Next i
  
ws.Range("K:K").Style = "Percent"
ws.Columns("I:L").AutoFit

Next ws

End Sub


