Sub getTickerAndVolume()
    'Reset excel functions before closing 
    Call performanceReset
    'Call function to get Ticker name and sum Volume
   Call allWorksheets
End Sub

'Call every page in Workbook
Function allWorksheets()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Call getTickerVolume
    Next
    MsgBox("Script completed")
End Function

'Fetching Tickers and calculating Total Stock Volume
Function getTickerVolume()
    'Shut down temporaly excel functions to improve script performance
    Call performanceBoost
    'Retrieve names of Ticker 1 and 2
    Dim Ticker1, Ticker2 As String
    'Retrieve current and total Volume
    Dim currentVolume, totalVolume As Double
    'To control location where Ticker and Volume will be stored
    Dim currentRow As Integer
    'To find the opening value of a Ticker
    Dim loopCounter As Integer
    'Last row in every WorkSheet
    Dim lastRow As Long
    'Initialize variables
    totalVolume = 0
    currentRow = 1
    loopCounter = -1
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Loop to retrieve Tickers, compare values, sum volumes and place info
    For i = 2 To lastRow
        Ticker1 = Cells(i, 1).Value2
        Ticker2 = Cells(i + 1, 1).Value2
        currentVolume = Cells(i, 7).Value2
        totalVolume = totalVolume + currentVolume
        loopCounter = loopCounter + 1
        
        If (Ticker2 <> Ticker1) Then
            currentRow = currentRow + 1
            Cells(currentRow, 9).Value = Ticker1
            Cells(currentRow, 12).Value = totalVolume
            Call getTickerStockVolume(currentRow, i, loopCounter)
            'Reset variables
            totalVolume = 0
            loopCounter = -1
        End If
    Next i
    tickerGreatest(currentRow)
    'Active Excel funcion
    Call performanceReset
End Function

'Calculating and fomatting yearly Change column
Function getTickerStockVolume(n_currentRow, n_i, n_loopCounter)
    'Variables for Open & Close price
    Dim openPrice, closePrice As Double
    'Retrieve Open and Close prices to get Yearly Change
    openPrice = Cells(n_i - n_loopCounter, 3).Value2
    closePrice = Cells(n_i, 6).Value2
    Cells(n_currentRow, 10).Value = closePrice - openPrice
    Cells(n_currentRow, 10).NumberFormat = "0.000000000"
    'Formating conditional according Yearly change value
    If (Cells(n_currentRow, 10).Value) >= 0 Then
        Cells(n_currentRow, 10).Interior.ColorIndex = 4
    Else
        Cells(n_currentRow, 10).Interior.ColorIndex = 3
    End If
    'Get Percent change
    If (openPrice = 0 Or closePrice = 0) Then
        Cells(n_currentRow, 11).Value = 0
    Else
        Cells(n_currentRow, 11).Value = closePrice / openPrice - 1
    End If
    Cells(n_currentRow, 11).NumberFormat = "0.00%"
End Function

'Calculating greatest
Function tickerGreatest(countTotalRows)
    Dim tickerRows As Integer
    Dim currentPercent, maxPercent, currentStockVolume, maxStockVolume As Double
    tickerRows = countTotalRows
    'Initialize variables
    maxPercent = 0
    minPercent = 0
    maxStockVolume = 0
    'Get Max/Min Percent Change and Total Stock Volume
    For i = 2 To tickerRows
        'Get Max Increase
        currentPercent = Cells(i, 11).Value2
        If currentPercent > maxPercent Then
            maxPercent = currentPercent
            Range("P2").Value = Cells(i, 9).Value2
            Range("Q2").Value = maxPercent
            Range("Q2").NumberFormat = "0.00%"
        End If
        'Get Min Decrease
        If currentPercent < minPercent Then
            minPercent = currentPercent
            Range("P3").Value = Cells(i, 9).Value2
            Range("Q3").Value = minPercent
            Range("Q3").NumberFormat = "0.00%"
        End If
        'Get Max Total Stock Volume
        currentStockVolume = Cells(i, 12).Value2
        If currentStockVolume > maxStockVolume Then
            maxStockVolume = currentStockVolume
            Range("P4").Value = Cells(i, 9).Value2
            Range("Q4").Value = maxStockVolume
        End If

    Next i
    'Funtion to build labels for new data
    Call buildColumnLabels
End Function

Function buildColumnLabels()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greates % Increase"
    Cells(3, 15).Value = "Greates % Decrease"
    Cells(4, 15).Value = "Greates Total Volume"
End Function

'A couple function to improve script performance
Function performanceBoost()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Function

Function performanceReset()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function