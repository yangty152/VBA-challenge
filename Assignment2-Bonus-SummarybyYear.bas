Attribute VB_Name = "Module3"
Sub SummaryByYear()

'This program will look at the summary by ticker table, and find the stock ticker with the greatest % increase, and the the greatest % decrease, and the greatest total stock volumn
'Loop through all worksheet
For Each ws In Worksheets
    'Declare variables to mark last row of the summary table, and hold temp values while to look through the data set to find the greatest values.
    Dim lastrow As Integer
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolumn As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumnTicker As String
    
    'Inititate variables
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    MaxIncrease = ws.Cells(2, 11).Value
    MaxDecrease = ws.Cells(2, 11).Value
    MaxVolumn = ws.Cells(2, 12).Value
    
    'Check each row in the sumamry table, and compare current row with previous row to find the greatest values
    For i = 2 To lastrow
        If MaxIncrease > ws.Cells(i + 1, 11).Value Then
            MaxIncrease = MaxIncrease
        Else
            MaxIncrease = ws.Cells(i + 1, 11).Value
        End If
        If MaxDecrease > ws.Cells(i + 1, 11).Value Then
            MaxDecrease = ws.Cells(i + 1, 11).Value
        Else
            MaxDecrease = MaxDecrease
        End If
        If MaxVolumn > ws.Cells(i + 1, 12).Value Then
            MaxVolumn = MaxVolumn
        Else
            MaxVolumn = ws.Cells(i + 1, 12).Value
        End If
    Next i
    
    'Check each row in the summary table to find the ticker for the greatest values idendified in the previous step.
    For j = 2 To lastrow
        If MaxIncrease = ws.Cells(j, 11) Then MaxIncreaseTicker = ws.Cells(j, 9)
        If MaxDecrease = ws.Cells(j, 11) Then MaxDecreaseTicker = ws.Cells(j, 9)
        If MaxVolumn = ws.Cells(j, 12) Then MaxVolumnTicker = ws.Cells(j, 9)
    Next j
    
    'Assign values to the summary by year table
    ws.Cells(2, 15).Value = MaxIncreaseTicker
    ws.Cells(3, 15).Value = MaxDecreaseTicker
    ws.Cells(4, 15).Value = MaxVolumnTicker
    ws.Cells(2, 16).Value = MaxIncrease
    ws.Cells(3, 16).Value = MaxDecrease
    ws.Cells(4, 16).Value = MaxVolumn
    
Next ws
End Sub
