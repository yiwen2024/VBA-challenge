Attribute VB_Name = "Module1"
Sub YearlyStock()

    Dim ws As Worksheet

    For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    Dim Ticker As String
    Ticker = " "
    
    Dim Ticker_volume As Double
    Ticker_volume = 0

    Dim stock_volume As LongLong
    stock_volume = 0

    Dim Lastrow As Long
    Dim i As Long

    Dim open_price As Double
    open_price = 0
    
    Dim close_price As Double
    close_price = 0
    
    Dim price_change As Double
    price_change = 0
    
    Dim Percent_Change As Double
    Percent_Change = 0
    
    Dim Yearly_Change As Double
    Yealy_Change = 0
    
    Dim TickerRow As Long
    TickerRow = 1
    
    Dim first_row As Long
    first_row = 2
    
    Dim MaxChange As Double
    MaxChange = 0
    
    Dim MinChange As Double
    MinChange = 0
    
    Dim MaxChangeTicker As String
    MaxChangeTicker = ""
    
    Dim MinChangeTicker As String
    MinChangeTicker = ""
    
   
    Dim MaxVolume As LongLong
    MaxVolume = 0
    
    Dim MaxVolumeTicker As String
    MaxVolumeTicker = ""
    
    For i = 2 To Rows.Count - 1
    
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          TickerRow = TickerRow + 1
          
          Ticker = ws.Cells(i, 1).Value
          ws.Cells(TickerRow, "I").Value = Ticker
          
          close_price = ws.Cells(i, 6).Value
          open_price = ws.Cells(first_row, 3).Value
          
          Ticker_volume = ws.Cells(i, 7).Value
          stock_volume = stock_volume + Ticker_volume
          
          Yealy_Change = close_price - open_price
          Percent_Change = (Yealy_Change / open_price)
          ws.Cells(TickerRow, "K").NumberFormat = "0.00%"
          
          ws.Cells(TickerRow, "J").Value = Yealy_Change
          ws.Cells(TickerRow, "K").Value = Percent_Change
          ws.Cells(TickerRow, "L").Value = stock_volume
       
          If Percent_Change > MaxChange Then
             MaxChange = Percent_Change
             MaxChangeTicker = Ticker
          End If
          
          If Percent_Change < MinChange Then
             MinChange = Percent_Change
             MinChangeTicker = Ticker
          End If
          
          If stock_volume > MaxVolume Then
             MaxVolume = stock_volume
             MaxVolumeTicker = Ticker
          End If
          
          If ws.Cells(TickerRow, "J").Value < 0 Then
             ws.Cells(TickerRow, "J").Interior.Color = vbRed
          ElseIf ws.Cells(TickerRow, "J").Value > 0 Then
             ws.Cells(TickerRow, "J").Interior.Color = vbGreen
          End If
      
          first_row = i + 1
          stock_volume = 0
      Else
          Ticker_volume = ws.Cells(i, 7).Value
          stock_volume = stock_volume + Ticker_volume
      
      End If
       
    Next i
    
    ws.Cells(2, 17) = MaxChange
    ws.Cells(2, 16) = MaxChangeTicker
    
    ws.Cells(3, 17) = MinChange
    ws.Cells(3, 16) = MinChangeTicker
    
    ws.Cells(4, 17) = MaxVolume
    ws.Cells(4, 16) = MaxVolumeTicker
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    Next ws

End Sub


