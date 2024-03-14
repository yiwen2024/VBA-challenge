# VBA-challenge
README
This module 2 challenge is to practice using VBA scripting to analyze generated stock market data. The goals to achieve are: 
       >The ticker symbol
       >Yearly change from the opening price at the beginning of a given year to the closing price at the 
        end of that year.
       >The percentage change from the opening price at the beginning of a given year to the closing     
        price at the end of that year.
       >The total stock volume of the stock.
       >Add functionality to your script to return the stock with the "Greatest % increase", "Greatest %  
        decrease", and "Greatest total volume". 
       >Make the appropriate adjustments to your VBA script to enable it to run on every worksheet          
        at once

The assignment was performed by referring to various online resources including Google, Stack Overflow, Reddit, etc. I appreciate all the shared information and I put those pieces of information together and did numerous trials before making the code work. Through practice, I better understood the format and logic including variables and coding conditions. Also, I learned different ways to define variables. For instance, both Cells (i, 1). Value" and Range ("A").Value can work. The coding with an explanation is shown below: 

Sub YearlyStock()

'Define ws to apply one code to all three worksheets

    Dim ws As Worksheet

    For Each ws In Worksheets

'Define the location of variables in the cells 
 
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

'Define variables for the calculation 

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

'Loop corresponding rows and columns for calculation. The logic is that the stock values with the same Ticker name will be added up until the Ticker name is changed in the next row. 
'The definition of fist_row is the open_price of the Ticker in a year in order to calculate the Year_Change through being subtracted by the close_price.
  
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
          percen_change = (Yealy_Change / open_price) * 100
          ws.Cells(TickerRow, "K").NumberFormat = "0.00%"
          
          ws.Cells(TickerRow, "J").Value = Yealy_Change
          ws.Cells(TickerRow, "K").Value = percen_change
          ws.Cells(TickerRow, "L").Value = stock_volume
      
          first_row = i + 1
          stock_volume = 0
      Else
          Ticker_volume = ws.Cells(i, 7).Value
          stock_volume = stock_volume + Ticker_volume

 'Define the variables for the summary session to extract the maximum change, minimum change, and maximum sock volume and corresponding Ticker name. 
 'Another way to find the max value is Application.WorksheetFunction.Max(Range).
        
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
 
'To color the interior cells as required. Different color index can be used including RBG (255, 0, 0) for red. 
      
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
 
'To define the location of the extreme value after calculation.
   
    ws.Cells(2, 17) = MaxChange
    ws.Cells(2, 16) = MaxChangeTicker
    
    ws.Cells(3, 17) = MinChange
    ws.Cells(3, 16) = MinChangeTicker
    
    ws.Cells(4, 17) = MaxVolume
    ws.Cells(4, 16) = MaxVolumeTicker

'To give the format of the percentage values.
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"

'To apply the same code to another worksheet.
    
    Next ws

End Sub

   
