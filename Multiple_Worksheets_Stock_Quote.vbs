Sub stock_data()

  ' --------------------------------------------
  ' LOOP THROUGH ALL SHEETS
  ' --------------------------------------------
  For Each ws In Worksheets

  '
  ' Define output headers for yearly change, percent change and total volume
  '
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
  ' Define working variables to derive these fields for all stocks
  Dim Ovalue, Cvalue, Volume As Double
  Dim Tcount, StockNum As Long
  Volume = 0
  Tcount = 0
  StockNum = 1
  '
  ' Define three most right columns aggregating greatest increase & decrease
  '
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
    
  ' Define ticker string variables for for yearly increase, yearly decrease and stock valume
  Dim GIticker, GDticker, GTticker As String
  
  ' Set float and integer variables for yearly increase, yearly decrease and stock valume
  Dim GIvalue, GDvalue, GTvolume As Double
  GIvalue = 0
  GDvalue = 0
  GTvolume = 0
  
  '
  ' Loop through all stocks contained in a single worksheet
  Dim LR As Long
  ' or understanding LR = Last Row

  LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

  For I = 2 To LR

    ' Check if we are still processing the same stock, if not ...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Increase the number of unique stocks analyzed
      StockNum = StockNum + 1
      
      ' Set the Close Value
      Cvalue = ws.Cells(I, 6).Value
        
      ' Add to the Volume
      Volume = Volume + ws.Cells(I, 7).Value

      ' Populate now the ticker, yearly change, percent change and total volume
      ws.Cells(StockNum, 9).Value = ws.Cells(I, 1).Value
      ws.Cells(StockNum, 10).Value = Cvalue - Ovalue
      If ws.Cells(StockNum, 10).Value > 0 Then
        ' Color the Yearly Change green
        ws.Cells(StockNum, 10).Interior.ColorIndex = 4
      Else
        ' Color the Yearly Change redColor the Yearly Change red
        ws.Cells(StockNum, 10).Interior.ColorIndex = 3
      End If
      
      ' Make sure Ovalue is non Zero for deriving the percentage value
      If Ovalue > 0 Then
        ws.Cells(StockNum, 11).Value = ws.Cells(2, 10).Value / Ovalue * 100
        ws.Cells(StockNum, 12).Value = Volume
      End If

      ' Update the values of greatest increase, decrease and total volume
      If ws.Cells(StockNum, 10).Value > GIvalue Then
        GIvalue = ws.Cells(StockNum, 10).Value
        GIticker = ws.Cells(StockNum, 9).Value
      End If
      
      If ws.Cells(StockNum, 10).Value < GDvalue Then
        GDvalue = ws.Cells(StockNum, 10).Value
        GDticker = ws.Cells(StockNum, 9).Value
      End If
     
      If ws.Cells(StockNum, 12).Value > GTvolume Then
        GTvolume = ws.Cells(StockNum, 12).Value
        GTticker = ws.Cells(StockNum, 9).Value
      End If
      
      ' Reset the Volume and Tcount
      Volume = 0
      Tcount = 0

    ' If the cell immediately following a row is the same stock ...
    Else

      ' Set the Open value
      If Tcount = 0 Then
        Ovalue = ws.Cells(I, 3).Value
      End If
      
      ' Add to the Volume and Tcount
      Volume = Volume + ws.Cells(I, 7).Value
      Tcount = Tcount + 1

    End If

  Next I
  
  ' Fill values for Greatest Intrease, Decrease and Volume data
  ws.Cells(2, 16).Value = GIticker
  ws.Cells(2, 17).Value = GIvalue
  ws.Cells(3, 16).Value = GDticker
  ws.Cells(3, 17).Value = GDvalue
  ws.Cells(4, 16).Value = GTticker
  ws.Cells(4, 17).Value = GTvolume
  
  
  ' --------------------------------------------
  ' FIXES COMPLETE
  ' --------------------------------------------

  Next ws

  MsgBox ("All Worksheets Processed. Exiting!!!")

End Sub
