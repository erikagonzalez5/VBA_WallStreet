Sub MultiYearStockData():

 'Setting the Variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearly As Double
    Dim percent As Double
    Dim volume As Double
    Dim i As Long
    Dim tickerCount As Integer
    
'worksheets and last row

    For Each ws In Worksheets
    ws.Activate

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'starting counts
    volume = 0
    tickerCount = 2


'ticker loop
    For i = 2 To LastRow

'setting open price
    openPrice = Cells(i, 3).Value

'setting close price
    closePrice = Cells(i, 6).Value

' setting total volume
    volume = volume + Cells(i, 7).Value

    'tickers
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            Cells(tickerCount, 9) = ticker


'calculating the yearly price change
    yearly = closePrice - openPrice

    'color coding the yearly price change column
         Cells(tickerCount, 10).Value = yearly
             If Cells(tickerCount, 10).Value > = 0 Then
             Cells(tickerCount, 10).Interior.ColorIndex = 4

             Else
             Cells(tickerCount, 10).Interior.ColorIndex = 3

            End If

    'making sure Yearly Change Header has a white background
        Range("J1").Interior.ColorIndex = 0

    'percent change calculation
             If (openPrice = 0 And closePrice = 0) Then
            percent = 0

             Else
             percent = yearly / openPrice
        End If


    'change to percent
        Cells(tickerCount, 11).Value = Format(percent, "Percent")


    'stock volume
        openPrice = 0

    'setting volume to the correct area within table
        Cells(tickerCount, 12).Value = volume

        volume = 0
    'row + 1 iteration (know's to move to the next row)
        tickerCount = tickerCount + 1

        End If

    Next i
    
Next ws

End Sub



