Attribute VB_Name = "QuaterlyStockSummary_Q1"
Sub QuarterlyStockSummary_Q1()

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Add headers for the summary table (Columns I-L)
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    Dim currentTicker As String
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim outputRow As Long
    
    outputRow = 2  ' Summary output starts at row 2
    
    ' Loop through each data row (assumes headers in row 1)
    For i = 2 To lastRow
        currentTicker = ws.Cells(i, 1).Value
        
        ' Process only rows within Q1 (January 2, 2022 to March 31, 2022)
        If ws.Cells(i, 2).Value >= DateValue("1/2/2022") And ws.Cells(i, 2).Value <= DateValue("3/31/2022") Then
            ' For the first Q1 row for a ticker, record the opening price (Column C)
            If openingPrice = 0 Then
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Update the closing price (Column F) and accumulate volume (Column G)
            closingPrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
        
        ' Check if the next row is a new ticker or if we've reached the last row
        Dim nextTicker As String
        If i < lastRow Then
            nextTicker = ws.Cells(i + 1, 1).Value
        Else
            nextTicker = ""
        End If
        
        ' When the ticker group ends, calculate and output the summary for the current ticker
        If i = lastRow Or nextTicker <> currentTicker Then
            If openingPrice <> 0 Then
                quarterlyChange = closingPrice - openingPrice
                ' Calculate percentage change as a fraction (e.g., 0.15 for 15%)
                If openingPrice <> 0 Then
                    percentageChange = quarterlyChange / openingPrice
                Else
                    percentageChange = 0
                End If
                
                ' Output summary results in columns I (Ticker), J (Quarterly Change),
                ' K (Percent Change) and L (Total Stock Volume)
                ws.Cells(outputRow, 9).Value = currentTicker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Apply color coding for Quarterly Change: green for positive, red for negative
                If quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(144, 238, 144)
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 182, 193)
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = xlNone
                End If
                
                outputRow = outputRow + 1
            End If
            
            ' Reset values for the next ticker group
            openingPrice = 0
            totalVolume = 0
        End If
    Next i
    
    ' --- Additional Functionality ---
    ' Scan the summary table (rows 2 to outputRow-1) to determine:
    ' 1. The ticker with the greatest % increase (max value in Column K)
    ' 2. The ticker with the greatest % decrease (min value in Column K)
    ' 3. The ticker with the greatest total volume (max value in Column L)
    
    If outputRow > 2 Then
        Dim maxPct As Double, minPct As Double, maxVol As Double
        Dim maxPctTicker As String, minPctTicker As String, maxVolTicker As String
        Dim summaryRow As Long
        
        ' Initialize using the first summary row
        maxPct = ws.Cells(2, 11).Value
        minPct = ws.Cells(2, 11).Value
        maxVol = ws.Cells(2, 12).Value
        maxPctTicker = ws.Cells(2, 9).Value
        minPctTicker = ws.Cells(2, 9).Value
        maxVolTicker = ws.Cells(2, 9).Value
        
        For summaryRow = 2 To outputRow - 1
            If ws.Cells(summaryRow, 11).Value > maxPct Then
                maxPct = ws.Cells(summaryRow, 11).Value
                maxPctTicker = ws.Cells(summaryRow, 9).Value
            End If
            If ws.Cells(summaryRow, 11).Value < minPct Then
                minPct = ws.Cells(summaryRow, 11).Value
                minPctTicker = ws.Cells(summaryRow, 9).Value
            End If
            If ws.Cells(summaryRow, 12).Value > maxVol Then
                maxVol = ws.Cells(summaryRow, 12).Value
                maxVolTicker = ws.Cells(summaryRow, 9).Value
            End If
        Next summaryRow
        
        ' Add headers for the additional output:
        ' Column P header: "Ticker" and Column Q header: "Value"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Output the additional results starting in row 2 (Columns O, P, Q)
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxPctTicker
        ws.Cells(2, 17).Value = maxPct
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = minPctTicker
        ws.Cells(3, 17).Value = minPct
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = maxVolTicker
        ws.Cells(4, 17).Value = maxVol
    End If

End Sub
