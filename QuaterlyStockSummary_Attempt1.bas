Attribute VB_Name = "QuaterlyStockSummary_Attempt1"
Sub QuarterlyStockSummary()

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ' Add headers to columns 9-12
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Last row in the data (assuming the data is continuous)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the rows of data
    Dim i As Long
    Dim ticker As String
    Dim startQuarterRow As Long
    Dim endQuarterRow As Long
    Dim totalVolume As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    
    ' Assuming columns based on the clarification:
    ' A: Ticker, B: Date, C: Open Price, F: Close Price, G: Volume
    
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value ' Ticker symbol
        
        ' Check if the date is within Q1 (adjust this logic for other quarters as needed)
        If ws.Cells(i, 2).Value >= DateValue("1/2/2022") And ws.Cells(i, 2).Value <= DateValue("3/31/2022") Then
            
            ' Find first opening price and last closing price in the quarter
            If openingPrice = 0 Then
                openingPrice = ws.Cells(i, 3).Value ' Opening price (Column 3)
                startQuarterRow = i ' Remember where the quarter starts
            End If
            
            closingPrice = ws.Cells(i, 6).Value ' Closing price (Column 6)
            totalVolume = totalVolume + ws.Cells(i, 7).Value ' Summing volume (Column 7)
            endQuarterRow = i ' Remember where the quarter ends
        End If
        
        ' If we reach the end of the quarter or the end of the data
        If ws.Cells(i + 1, 2).Value > DateValue("3/31/2022") Or ws.Cells(i + 1, 1).Value <> ticker Then
            ' Calculate changes if we are at the end of the quarter or ticker changes
            If openingPrice <> 0 Then
                quarterlyChange = closingPrice - openingPrice
                percentageChange = ((closingPrice - openingPrice) / openingPrice) * 100
                
                ' Output results to columns 9-12 (Columns I-L)
                ws.Cells(endQuarterRow, 9).Value = ticker ' Column I: Ticker
                ws.Cells(endQuarterRow, 10).Value = quarterlyChange ' Column J: Quarterly Change
                ws.Cells(endQuarterRow, 11).Value = percentageChange ' Column K: Percentage Change
                ws.Cells(endQuarterRow, 12).Value = totalVolume ' Column L: Total Volume
                
                ' Apply color coding to the Quarterly Change column (Column 10)
                If quarterlyChange > 0 Then
                    ws.Cells(endQuarterRow, 10).Interior.Color = RGB(144, 238, 144) ' Light green for positive
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(endQuarterRow, 10).Interior.Color = RGB(255, 182, 193) ' Light red for negative
                Else
                    ws.Cells(endQuarterRow, 10).Interior.ColorIndex = xlNone ' No color for 0
                End If
                
                ' Reset values for the next quarter or ticker
                openingPrice = 0
                totalVolume = 0
            End If
        End If
    Next i

End Sub


