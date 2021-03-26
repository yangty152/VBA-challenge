Attribute VB_Name = "Module1"
Sub SummarybyTicker()

'This program will return ticker symbol, yearly change (closeing price - opening price), percentage change(yearly change/openning price), total stock volumn per ticker.
'Data is orgnized by year, so the program will loop through each year (sheet) one by one.
For Each ws In Worksheets

'Declare variables
'Summary_Table_Row_Num will store the row number for the summary table, starting value as 2
'TickerFirstRowNum will store the first row of each ticker location, starting value as 2
'lastrow will store last row number of each sheet for the original dataset

    Dim Summary_Table_Row_Num As Integer
    Dim TickerFirstRowNum As Double
    Dim YearlyChange As Double
    Dim TotalVolumn As Double
    Dim lastrow As Double
   
 'Iinitiate values for the variables
    Summary_Table_Row_Num = 2
    TickerFirstRowNum = 2
    TotalVolumn = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'For each row in the original data set beside the header, find the unique ticker, and calcuate yearly change, percent change and total stock volumn
    For i = 2 To lastrow
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
        'Assign value to Ticker column
        ws.Cells(Summary_Table_Row_Num, 9).Value = ws.Cells(i, 1).Value
        'Calculate TotalVolumn for the last row of one ticker
        TotalVolumn = TotalVolumn + ws.Cells(i, 7).Value
        'Assign value to Total Stock Volumn column
        ws.Cells(Summary_Table_Row_Num, 12) = TotalVolumn
        'Calculate Yearly Change for the stock, compare the opennng price to the close price of the year for one ticker.
        YearlyChange = ws.Cells(i, 6).Value - ws.Cells(TickerFirstRowNum, 3).Value
        'Assign value to the yearly change column
        ws.Cells(Summary_Table_Row_Num, 10).Value = YearlyChange
        'conditioanl formatting when yearly change less or equal to 0, color it red, and when it is greater than 0, color it green
        If YearlyChange <= 0 Then
        ws.Cells(Summary_Table_Row_Num, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(Summary_Table_Row_Num, 10).Interior.ColorIndex = 4
        End If
        'Assign value to the percent change column, handle situation when the openning price is 0
        If ws.Cells(TickerFirstRowNum, 3).Value = 0 Then
            ws.Cells(Summary_Table_Row_Num, 11).Value = 0
        Else
            ws.Cells(Summary_Table_Row_Num, 11).Value = YearlyChange / ws.Cells(TickerFirstRowNum, 3).Value
        End If
        'Reset TotalVolumn for the next ticker
        TotalVolumn = 0
        'Move the summary table to the next row for the new ticker
        Summary_Table_Row_Num = Summary_Table_Row_Num + 1
        'Mark the first row index of a new ticker for the original data set
        TickerFirstRowNum = i + 1
    Else
        'If next row ticker is the same as current row, then only need to add the total to the volumn
        TotalVolumn = TotalVolumn + ws.Cells(i, 7).Value
    End If
    Next i
    'Format the percentage change column to percent format.
    ws.Range("K2:K" & Summary_Table_Row_Num).NumberFormat = "0.00%"
Next ws
End Sub

