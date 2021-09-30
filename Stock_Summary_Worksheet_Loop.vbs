Sub allWorksheetsStockSummary()

'loop through all sheets
For Each ws In Worksheets
' Set data types for variables
Dim tickerSymbol As String
Dim yearChange As Double
Dim percentChange As Double
Dim totalStockVolume As Double
Dim summaryTableRow As Integer

' set headers of summary table
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

' set boundaries of For loop
summaryTableRow = 2
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
    ' when ticker doesn't match the ticker from the following row
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerSymbol = ws.Cells(i, 1).Value
        'print ticker symbol in summary table
        ws.Range("I" & summaryTableRow).Value = tickerSymbol
        
        'save closing price to memory
        closingPriceEndYear = ws.Range("F" & i)
        'save opening price to memory
        openingPriceRow = i - rowCounter
        openingPriceBegYear = ws.Range("c" & openingPriceRow)
        'calculate year change in price
        yearChange = closingPriceEndYear - openingPriceBegYear
        'print yearly change in price to summary table
        ws.Range("J" & summaryTableRow).Value = yearChange
                        
        'Error handling for division by zero values
        If openingPriceBegYear = 0 Then
            percentChange = 0
        Else
        'calculate percent change in price throughout the year
            percentChange = yearChange / openingPriceBegYear
        End If
        'print percent change to the summary table
        ws.Range("K" & summaryTableRow).Value = percentChange
        'reset counter to 0 for next set of ticker symbols
        rowCounter = 0
        
        'add last row of stock volume to the total
        totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
        'print total stock volume to summary table
        ws.Range("L" & summaryTableRow).Value = totalStockVolume
        'reset total of stock Volume for next set of ticker symbols
        totalStockVolume = 0
        
        'move down to the next row in the summary table
        summaryTableRow = summaryTableRow + 1
    Else
        'when ticker symbols match, add this row's stock volume to
        'the running volume total and add one to the number of rows with this ticker symbol
        totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
        rowCounter = rowCounter + 1
    End If
Next i

'Formatting the summary table

'Get the number of rows in the summary table
lastRowSummaryTable = ws.Cells(Rows.Count, 10).End(xlUp).Row
For i = 2 To lastRowSummaryTable
    'when Yearly change is negative, highlight cell red
    If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    Else
    'Otherwise highlight cell green
        ws.Cells(i, 10).Interior.ColorIndex = 4
    End If
Next i
    'format percent change column to percentage format
    ws.Columns("K").NumberFormat = "0.00%"

'Create headers for bonus table
ws.Range("P1") = "Ticker"
ws.Range("q1") = "Value"
ws.Range("o2") = "Greatest % Increase"
ws.Range("o3") = "Greatest % Decrease"
ws.Range("o4") = "Greatest Total Volume"
'Initiate variables with first row in summary table's data
MaxIncrease = ws.Cells(2, 11).Value
MaxTicker = ws.Cells(2, 9).Value
decrease = ws.Cells(2, 11).Value
DecreaseTicker = ws.Cells(2, 9).Value
volume = ws.Cells(2, 12).Value
volumeticker = ws.Cells(2, 9).Value

'loop through entire summary table to find max increase, max decrease, and max total volume
For j = 2 To lastRowSummaryTable
        If ws.Cells(j + 1, 11).Value > MaxIncrease Then
            MaxIncrease = ws.Cells(j + 1, 11).Value
            MaxTicker = ws.Cells(j + 1, 9).Value
        Else
        End If
        
        If ws.Cells(j + 1, 11).Value < decrease Then
            decrease = ws.Cells(j + 1, 11).Value
            DecreaseTicker = ws.Cells(j + 1, 9).Value
        Else
        End If
        
        If ws.Cells(j + 1, 12).Value > volume Then
            volume = ws.Cells(j + 1, 12).Value
            volumeticker = ws.Cells(j + 1, 9).Value
        Else
        End If
Next j
'insert values in bonus table and reformat them to percent and scientific notation
    ws.Cells(2, 16).Value = MaxTicker
    ws.Cells(2, 17).Value = MaxIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = DecreaseTicker
    ws.Cells(3, 17).Value = decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = volumeticker
    ws.Cells(4, 17).Value = volume
    ws.Cells(4, 17).NumberFormat = "0.0000E+00"
   
Next ws
End Sub