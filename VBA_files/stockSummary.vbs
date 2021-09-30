Sub stockSummaryTable()

' Set data types for variables
Dim tickerSymbol As String
Dim yearChange As Double
Dim percentChange As Double
Dim totalStockVolume As Double
Dim summaryTableRow As Integer

' set headers of summary table
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

' set boundaries of For loop
summaryTableRow = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
    ' when ticker doesn't match the ticker from the following row
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        tickerSymbol = Cells(i, 1).Value
        'print ticker symbol in summary table
        Range("I" & summaryTableRow).Value = tickerSymbol
        
        'save closing price to memory
        closingPriceEndYear = Range("F" & i)
        'save opening price to memory
        openingPriceRow = i - rowCounter
        openingPriceBegYear = Range("c" & openingPriceRow)
        'calculate year change in price
        yearChange = closingPriceEndYear - Range("C" & openingPriceRow)
        'print yearly change in price to summary table
        Range("J" & summaryTableRow).Value = yearChange
                        
        'Error handling for division by zero values
        If openingPriceBegYear = 0 Then
            percentChange = 0
        Else
        'calculate percent change in price throughout the year
            percentChange = yearChange / openingPriceBegYear
        End If
        'print percent change to the summary table
        Range("K" & summaryTableRow).Value = percentChange
        'reset counter to 0 for next set of ticker symbols
        rowCounter = 0
        
        'add last row of stock volume to the total
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        'print total stock volume to summary table
        Range("L" & summaryTableRow).Value = totalStockVolume
        'reset total of stock Volume for next set of ticker symbols
        totalStockVolume = 0
        
        'move down to the next row in the summary table
        summaryTableRow = summaryTableRow + 1
    Else
        'when ticker symbols match, add this row's stock volume to
        'the running volume total and add one to the number of rows with this ticker symbol
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        rowCounter = rowCounter + 1
    End If
Next i

'Formatting the summary table

'Get the number of rows in the summary table
lastRowSummaryTable = Cells(Rows.Count, 10).End(xlUp).Row
For i = 2 To lastRowSummaryTable
    'when Yearly change is negative, highlight cell red
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    Else
    'Otherwise highlight cell green
        Cells(i, 10).Interior.ColorIndex = 4
    End If
Next i
    'format percent change column to percentage format
    Columns("K").NumberFormat = "0.00%"

'Create headers for bonus table
Range("P1") = "Ticker"
Range("q1") = "Value"
Range("o2") = "Greatest % Increase"
Range("o3") = "Greatest % Decrease"
Range("o4") = "Greatest Total Volume"
'Initiate variables with first row in summary table's data
MaxIncrease = Cells(2, 11).Value
MaxTicker = Cells(2, 9).Value
decrease = Cells(2, 11).Value
DecreaseTicker = Cells(2, 9).Value
volume = Cells(2, 12).Value
volumeticker = Cells(2, 9).Value

'loop through entire summary table to find max increase, max decrease, and max total volume
For j = 2 To lastRowSummaryTable
        If Cells(j + 1, 11).Value > MaxIncrease Then
            MaxIncrease = Cells(j + 1, 11).Value
            MaxTicker = Cells(j + 1, 9).Value
        Else
        End If
        
        If Cells(j + 1, 11).Value < decrease Then
            decrease = Cells(j + 1, 11).Value
            DecreaseTicker = Cells(j + 1, 9).Value
        Else
        End If
        
        If Cells(j + 1, 12).Value > volume Then
            volume = Cells(j + 1, 12).Value
            volumeticker = Cells(j + 1, 9).Value
        Else
        End If
Next j
'insert values in bonus table and reformat them to percent and scientific notation
    Cells(2, 16).Value = MaxTicker
    Cells(2, 17).Value = MaxIncrease
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = DecreaseTicker
    Cells(3, 17).Value = decrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = volumeticker
    Cells(4, 17).Value = volume
    Cells(4, 17).NumberFormat = "0.0000E+00"
   
End Sub