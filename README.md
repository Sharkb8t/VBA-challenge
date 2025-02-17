VBA-challenge

Hi! This README file contains the VBScript code that I used to summarize quarterly stock data statistics for the year of 2022.

I ended with this script as it was the most optimized way that I found to summarize the statistical changes in stock data for each datasheet from the assigned Macro-Enabled Excel workbook.

I needed the script to both show the quarterly change in stock value as well as the percent change and the stock volume traded each quarter.

In addition I needed the script to also show a statistic for 'Greatest % Increase', 'Greatest % Decrease', and 'Greatest Total Volume' for each quarters statistics.

    Sub QuarterlyStockSummary_All()
        Dim ws As Worksheet
        Dim dateStart As Date, dateEnd As Date
        Dim lastRow As Long, outputRow As Long
        Dim i As Long, summaryRow As Long
        Dim currentTicker As String, nextTicker As String
        Dim totalVolume As Double, openingPrice As Double, closingPrice As Double
        Dim quarterlyChange As Double, percentageChange As Double
        Dim maxPct As Double, minPct As Double, maxVol As Double
        Dim maxPctTicker As String, minPctTicker As String, maxVolTicker As String

        ' Loop through each worksheet in the workbook
        For Each ws In ThisWorkbook.Worksheets
            ' Process only the worksheets named Q1, Q2, Q3, or Q4
            If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            
                ' Set the date range based on the worksheet name
                Select Case ws.Name
                    Case "Q1"
                        dateStart = DateValue("1/2/2022")
                        dateEnd = DateValue("3/31/2022")
                    Case "Q2"
                        dateStart = DateValue("4/1/2022")
                        dateEnd = DateValue("6/30/2022")
                    Case "Q3"
                        dateStart = DateValue("7/1/2022")
                        dateEnd = DateValue("9/30/2022")
                    Case "Q4"
                        dateStart = DateValue("10/1/2022")
                        dateEnd = DateValue("12/31/2022")
                End Select
            
                ' Clear any previous summary headers (optional)
                ws.Range("I1:L1").Clear
                ws.Range("O1:Q4").Clear
            
                ' Add headers for the summary table (Columns I-L)
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Quarterly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
            
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                outputRow = 2 ' Summary output starts at row 2
            
                ' Initialize variables for each worksheet
                openingPrice = 0
                totalVolume = 0
            
                ' Loop through each row in the data (assuming headers in row 1)
                For i = 2 To lastRow
                    currentTicker = ws.Cells(i, 1).Value
                
                    ' Process only rows within the specified date range
                    If ws.Cells(i, 2).Value >= dateStart And ws.Cells(i, 2).Value <= dateEnd Then
                        ' Record opening price on the first data row for the ticker
                        If openingPrice = 0 Then
                            openingPrice = ws.Cells(i, 3).Value
                        End If
                    
                        ' Update closing price and accumulate volume
                        closingPrice = ws.Cells(i, 6).Value
                        totalVolume = totalVolume + ws.Cells(i, 7).Value
                    End If
                
                    ' Determine if the next row starts a new ticker group or if we're at the last row
                    If i < lastRow Then
                        nextTicker = ws.Cells(i + 1, 1).Value
                    Else
                        nextTicker = ""
                    End If
                
                    ' When the ticker group ends, output the summary
                    If i = lastRow Or nextTicker <> currentTicker Then
                        If openingPrice <> 0 Then
                            quarterlyChange = closingPrice - openingPrice
                            If openingPrice <> 0 Then
                                percentageChange = quarterlyChange / openingPrice
                            Else
                                percentageChange = 0
                            End If
                        
                            ws.Cells(outputRow, 9).Value = currentTicker
                            ws.Cells(outputRow, 10).Value = quarterlyChange
                            ws.Cells(outputRow, 11).Value = percentageChange
                            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                            ws.Cells(outputRow, 12).Value = totalVolume
                        
                            ' Apply color coding for Quarterly Change
                            If quarterlyChange > 0 Then
                                ws.Cells(outputRow, 10).Interior.Color = RGB(144, 238, 144)  ' Light green
                            ElseIf quarterlyChange < 0 Then
                                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 182, 193)  ' Light red
                            Else
                                ws.Cells(outputRow, 10).Interior.ColorIndex = xlNone
                            End If
                        
                            outputRow = outputRow + 1
                        End If
                        ' Reset variables for the next ticker
                        openingPrice = 0
                        totalVolume = 0
                    End If
                Next i
            
                ' --- Additional Analysis ---
                ' Only proceed if there is at least one summary row
                If outputRow > 2 Then
                    ' Initialize with the first summary row
                    maxPct = ws.Cells(2, 11).Value
                    minPct = ws.Cells(2, 11).Value
                    maxVol = ws.Cells(2, 12).Value
                    maxPctTicker = ws.Cells(2, 9).Value
                    minPctTicker = ws.Cells(2, 9).Value
                    maxVolTicker = ws.Cells(2, 9).Value
                
                    ' Loop through the summary table (rows 2 to outputRow - 1)
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
                
                    ' Add headers for additional output in Columns P and Q
                    ws.Cells(1, 16).Value = "Ticker"
                    ws.Cells(1, 17).Value = "Value"
                
                    ' Output additional results in columns O (labels), P (tickers), and Q (values)
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
            End If
        Next ws
    End Sub

