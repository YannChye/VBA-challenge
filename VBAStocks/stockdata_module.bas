Attribute VB_Name = "stockdata_module"
Sub stockData()

'get number of worksheets as 'numSheets'
    Dim numSheets As Integer
    numSheets = ActiveWorkbook.Worksheets.Count
    
'get range of data for each sheet
    Dim lastRow() As Long, currSheet As Worksheet
    For i = 1 To numSheets
        Set currSheet = ActiveWorkbook.Worksheets(i)
        ReDim Preserve lastRow(i) 'lastRow(i) is an array listing the last row of of each worksheet
        lastRow(i) = currSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    Next i
    
'setting out the variables to be used later on
    Dim tickerUnique() As String, startrowUnique(), Index As Long 'variable for section 1, ref below
    Dim openPrice, closePrice, percentChange As Double, stockVol As Long 'variable for section 2
    Dim maxInc As Double, maxIncTicker As String 'variable for section 3
    Dim maxDec As Double, maxDecTicker As String 'variable for section 4
    Dim maxVol, maxVolTicker As String 'variable for section 5
    
'everything being done from here on is being looped through each worksheet
'get unique tickers (section 1)
    For i = 1 To numSheets
        Set currSheet = ActiveWorkbook.Worksheets(i)
        'resetting array and populating first row values
        ReDim tickerUnique(1) 'name of unique tickers
        tickerUnique(1) = currSheet.Cells(2, 1)
        ReDim startrowUnique(1)
        startrowUnique(1) = 2 'array that lists the starting row for each unique ticker, first ticker always starts on row 2
        Index = 1 'index/counter to count the number of unique tickers
        For j = 3 To lastRow(i)
            If tickerUnique(Index) <> currSheet.Cells(j, 1) Then 'sequentially comparing tickers to pick up change
                Index = Index + 1
                ReDim Preserve startrowUnique(Index) 'where there is a change in ticker, index increases by 1,
            'startrowUnique stores the row number of the next ticker
                startrowUnique(Index) = j
                ReDim Preserve tickerUnique(Index) 'tickerUnique stores the name of each unique ticker
                tickerUnique(Index) = currSheet.Cells(j, 1)
            End If
        Next j
           
 'writing out the headers
        currSheet.Range("I1") = "Ticker"
        currSheet.Range("J1") = "Yearly Change ($)"
        currSheet.Range("K1") = "Percent Change"
        currSheet.Range("L1") = "Total Stock Volume"
          
'writing tickers out in column 9
        For j = 1 To Index 'am aware that I'm reusing 'j', but as it has completed the previous procedure, should not affect the next one
            currSheet.Cells(1 + j, 9) = tickerUnique(j) 'always 1+ to account for header row
        Next j
    
'get yearly change (section 2)
'am aware that I'm only extracting open and close price from the first and last date, rather than reading each and every row
'but I find it more efficient this way. Also not embedding it withing the previous loops so that each section is clearer.
        For j = 1 To Index 'Index being output row number
            openPrice = currSheet.Cells(startrowUnique(j), 3) 'openPrice is the open price for the first day of the year
            'closePrice is the closing price for the last day of the year,
            'or in this case closing price that appeared one row before the opening price for the next ticker
            If j = Index Then
                closePrice = currSheet.Cells(lastRow(i), 6) 'closePrice for last row because can't compare to the next ticker
            Else
                closePrice = currSheet.Cells(startrowUnique(j + 1) - 1, 6)
            End If
            currSheet.Cells(1 + j, 10) = closePrice - openPrice 'writing yearly change out in column 10

'get percentage change
            If openPrice <> 0 Then 'avoiding division by zero error
                percentChange = (closePrice - openPrice) / openPrice
                currSheet.Cells(1 + j, 11).NumberFormat = "0.00%"
                currSheet.Cells(1 + j, 11) = percentChange
            Else
                currSheet.Cells(1 + j, 11) = "na" 'assuming I'm just putting in na if can't calculate percent change
            End If

'get totalstockvol, i.e. sum of column 7 for each ticker
            If j = Index Then
                currSheet.Cells(1 + j, 12) = WorksheetFunction.Sum(Range(currSheet.Cells(startrowUnique(j), 7), currSheet.Cells(lastRow(i), 7)))
            Else
                currSheet.Cells(1 + j, 12) = WorksheetFunction.Sum(Range(currSheet.Cells(startrowUnique(j), 7), currSheet.Cells(startrowUnique(1 + j) - 1, 7)))
            End If
        Next j
    'can also do alternate version to get totalstockvol (outside of the j for-loop)
        'RowCount = 1
        'totalstockvol = currSheet.Cells(2, 7)
        'For j = 2 To lastRow(i)
            'If currSheet.Cells(j, 1) = currSheet.Cells(j + 1, 1) Then
                'totalstockvol = totalstockvol + currSheet.Cells(j + 1, 7)
            'Else
                'RowCount = RowCount + 1
                'currSheet.Cells(RowCount, 12) = totalstockvol
                'totalstockvol = currSheet.Cells(j + 1, 7)
            'End If
        'Next j
    
'Colourcode the change
        For j = 1 To Index
            If currSheet.Cells(1 + j, 10) > 0 Then
                currSheet.Cells(1 + j, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf currSheet.Cells(1 + j, 10) < 0 Then  'assuming I'm just ignoring values that are exactly 0, therefore specifically not coloring in cell=0
                currSheet.Cells(1 + j, 10).Interior.Color = RGB(255, 0, 0)
            End If
        Next j
    
'Writing out the headers for challenge
        currSheet.Range("O:O").ColumnWidth = 25
        currSheet.Range("O2") = "Greatest % Increase"
        currSheet.Range("O3") = "Greatest % Decrease"
        currSheet.Range("O4") = "Greatest Total Volume"
        currSheet.Range("P1") = "Ticker"
        currSheet.Range("Q1") = "Value"

'Getting greatest % increase (section 3)
        maxInc = 0
        maxIncTicker = ""
        For j = 1 To Index
            If currSheet.Cells(1 + j, 11) = "na" Then 'ignoring where percent change could not be calculated
            ElseIf currSheet.Cells(1 + j, 11) > maxInc Then 'iteratively updating value only if next number is larger
                maxInc = currSheet.Cells(1 + j, 11)
                maxIncTicker = currSheet.Cells(1 + j, 9)
            End If
        Next j
        currSheet.Range("Q2").NumberFormat = "0.00%"
        currSheet.Range("P2") = maxIncTicker 'writing out values into worksheet
        currSheet.Range("Q2") = maxInc

'Getting greatest % decrease (section 4)
        maxDec = 0
        maxDecTicker = ""
        For j = 1 To Index
            If currSheet.Cells(1 + j, 11) = "na" Then
            ElseIf currSheet.Cells(1 + j, 11) < maxDec Then 'iteratively updating value only if next number is smaller
                maxDec = currSheet.Cells(1 + j, 11)
                maxDecTicker = currSheet.Cells(1 + j, 9)
            End If
        Next j
        currSheet.Range("Q3").NumberFormat = "0.00%"
        currSheet.Range("P3") = maxDecTicker 'writing out values into worksheet
        currSheet.Range("Q3") = maxDec

'Getting greatest total volume (section 5)
        maxVol = 0
        maxVolTicker = ""
        For j = 1 To Index
            If currSheet.Cells(1 + j, 12) > maxVol Then
                maxVol = currSheet.Cells(1 + j, 12)
                maxVolTicker = currSheet.Cells(1 + j, 9)
            End If
        Next j
        currSheet.Range("P4") = maxVolTicker
        currSheet.Range("Q4") = maxVol
    Next i

End Sub

