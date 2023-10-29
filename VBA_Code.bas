Attribute VB_Name = "Module1"

Sub stocksAnalysis()
    'Create a script that loops through all the stocks for one year and outputs the following information:
    'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
        'For each ws in Worksheets... Next ws
    
        
    For Each ws In Worksheets
    
    'Variables
    Dim tickerCode, previousTickerCode As String
    Dim previousTickerOpen, previousTickerClose, tickerClose, percentChange, totalStockVolume As Double
        previousTickerOpen = ws.Cells(2, 3)
        totalStockVolume = ws.Cells(2, 7)
    Dim tickerRow As Integer
        tickerRow = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim greatestIncrease, greatestDecrease As Integer
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker, greatestDecreaseTicker, greatestVolumeTicker As String
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
    'Headers
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(2, 14) = "Greatest % Increase"
    ws.Cells(3, 14) = "Greatest % Decrease"
    ws.Cells(4, 14) = "Greatest Total Volume"
    ws.Cells(1, 15) = "Ticker"
    ws.Cells(1, 16) = "Value"
        
    For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerCode = ws.Cells(i, 1).Value
        previousTickerClose = ws.Cells(i, 6)

        ws.Range("I" & tickerRow).Value = tickerCode
        ws.Range("L" & tickerRow) = totalStockVolume
           
        'Calculate the yearly change for the current ticker code and formats cells
        ws.Range("J" & tickerRow) = previousTickerClose - previousTickerOpen
            If ws.Range("J" & tickerRow) > 0 Then
                ws.Range("J" & tickerRow).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & tickerRow) < 0 Then
                ws.Range("J" & tickerRow).Interior.ColorIndex = 3
            End If
          
            'Summary table
            If ws.Range("J" & tickerRow) > greatestIncrease Then
                greatestIncrease = ws.Range("J" & tickerRow)
                greatestIncreaseTicker = ws.Range("I" & tickerRow).Value
            End If
            
            If ws.Range("J" & tickerRow).Value < greatestDecrease Then
                greatestDecrease = ws.Range("J" & tickerRow)
                greatestDecreaseTicker = ws.Range("I" & tickerRow).Value
            End If
            
            If ws.Range("L" & tickerRow).Value > greatestVolume Then
                greatestVolume = ws.Range("L" & tickerRow).Value
                greatestVolumeTicker = ws.Range("I" & tickerRow).Value
            End If
    
        'Calculate the percentage change for current ticker code and formats cells
        ws.Range("K" & tickerRow) = FormatPercent((ws.Range("J" & tickerRow) / previousTickerOpen), 2)
            If ws.Range("K" & tickerRow) > 0 Then
                    ws.Range("K" & tickerRow).Interior.ColorIndex = 4
            ElseIf ws.Range("K" & tickerRow) < 0 Then
                    ws.Range("K" & tickerRow).Interior.ColorIndex = 3
            End If
            
        
        tickerRow = tickerRow + 1
        totalStockVolume = ws.Cells(i + 1, 7)
        previousTickerOpen = ws.Cells(i + 1, 3)
    Else
        totalStockVolume = totalStockVolume + ws.Cells(i, 7)
            
    End If
    Next i
    
    
        'Print the values
        ws.Cells(2, 16) = greatestIncrease
        ws.Cells(2, 15) = greatestIncreaseTicker
        ws.Cells(3, 16) = greatestDecrease
        ws.Cells(3, 15) = greatestDecreaseTicker
        ws.Cells(4, 16) = greatestVolume
        ws.Cells(4, 15) = greatestVolumeTicker
    
    Next ws
    MsgBox ("Procedure complete")
    
End Sub
