VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis_KimNguyen()
    ' Loop through all worksheets
    For Each ws In Worksheets
        
        ' Define variables
        Dim tickerSymbol As String
        Dim quarterlyOpen As Double
        Dim quarterlyClose As Double
        Dim quarterlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        Dim summaryRow As Integer
        Dim lastRow As Long
        
        ' Variables for greatest values
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim increaseTickerSymbol As String
        Dim decreaseTickerSymbol As String
        Dim volumeTickerSymbol As String
        
        ' Initialize variables
        summaryRow = 2
        totalVolume = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        quarterlyOpen = ws.Cells(2, 3).Value  ' First opening price
        
        ' Create headers for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize greatest values
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if we're still within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set ticker symbol
                tickerSymbol = ws.Cells(i, 1).Value
                
                ' Get quarterly close price
                quarterlyClose = ws.Cells(i, 6).Value
                
                ' Calculate quarterly change
                quarterlyChange = quarterlyClose - quarterlyOpen
                
                ' Calculate percent change
                If quarterlyOpen <> 0 Then
                    percentChange = (quarterlyChange / quarterlyOpen)
                Else
                    percentChange = 0
                End If
                
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output results
                ws.Cells(summaryRow, 9).Value = tickerSymbol
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format percent change as percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Conditional formatting
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                End If
                
                 If percentChange >= 0 Then
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 43
                Else
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 53
                End If
        
                
                ' Check for greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerSymbol = tickerSymbol
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerSymbol = tickerSymbol
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerSymbol = tickerSymbol
                End If
                
                ' Reset variables for next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                quarterlyOpen = ws.Cells(i + 1, 3).Value
            Else
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Create greatest values summary
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ' Output greatest values
        ws.Cells(2, 15).Value = increaseTickerSymbol
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
        
        ws.Cells(3, 15).Value = decreaseTickerSymbol
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = volumeTickerSymbol
        ws.Cells(4, 16).Value = greatestVolume
        
        ' Autofit columns
        ws.Columns("I:P").AutoFit
    Next ws
End Sub
