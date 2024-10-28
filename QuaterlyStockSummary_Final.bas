Attribute VB_Name = "QuaterlyStockSummary_Final"
Sub SummarizeStockDataAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row with data in column A (ticker column) for the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        ' Define output start row and set headers for the current sheet
        outputRow = 2 ' First output row after headers
        ws.Range("I1:L1").Value = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")
        ws.Range("O1:Q1").Value = Array("", "Ticker", "Value")
        
        ' Collection to hold unique tickers
        Dim uniqueTickers As Object
        Set uniqueTickers = CreateObject("Scripting.Dictionary") ' Faster than Collection
        
        ' Loop through the tickers and collect unique values
        Dim i As Long, ticker As String
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            If Len(ticker) > 0 Then
                If Not uniqueTickers.exists(ticker) Then
                    uniqueTickers.Add ticker, ticker
                End If
            End If
        Next i
        
        ' Variables for greatest increase, decrease, and volume
        Dim greatestIncrease As Double: greatestIncrease = -9999999
        Dim greatestDecrease As Double: greatestDecrease = 9999999
        Dim greatestVolume As Double: greatestVolume = 0
        
        ' Variables to store associated tickers
        Dim increaseTicker As String, decreaseTicker As String, volumeTicker As String
        
        ' Loop through each unique ticker to process calculations
        Dim t As Variant, firstOpen As Double, lastClose As Double, totalVol As Double
        Dim quarterlyChange As Double, percentChange As Double
        Dim firstOpenFound As Boolean
        
        For Each t In uniqueTickers.Keys
            totalVol = 0
            firstOpenFound = False
            
            ' Loop through all rows to find first and last occurrence of ticker and sum volume
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value = t Then
                    ' Capture the first open value and the last close value
                    If Not firstOpenFound Then
                        firstOpen = ws.Cells(i, 3).Value
                        firstOpenFound = True
                    End If
                    lastClose = ws.Cells(i, 6).Value
                    totalVol = totalVol + ws.Cells(i, 7).Value
                End If
            Next i
            
            ' Calculate quarterly change and percent change
            quarterlyChange = lastClose - firstOpen
            percentChange = Round((quarterlyChange / firstOpen) * 100, 2) ' Percent change rounded to 2 decimals
            
            ' Write to output and apply formatting
            With ws.Cells(outputRow, 9)
                ' Write Ticker
                .Value = t
                ' Write Quarterly Change with number format and color formatting
                .Offset(0, 1).Value = quarterlyChange
                .Offset(0, 1).NumberFormat = "0.00"
                If quarterlyChange > 0 Then
                    .Offset(0, 1).Interior.Color = RGB(144, 238, 144) ' Green
                ElseIf quarterlyChange < 0 Then
                    .Offset(0, 1).Interior.Color = RGB(255, 99, 71) ' Red
                Else
                    .Offset(0, 1).Interior.ColorIndex = xlNone
                End If
                ' Write Percent Change with number format
                .Offset(0, 2).Value = percentChange
                .Offset(0, 2).NumberFormat = "0.00"
                ' Write Total Volume
                .Offset(0, 3).Value = totalVol
            End With
            
            ' Check for greatest increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                increaseTicker = t
            End If
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                decreaseTicker = t
            End If
            If totalVol > greatestVolume Then
                greatestVolume = totalVol
                volumeTicker = t
            End If
            
            ' Move to the next output row
            outputRow = outputRow + 1
        Next t
        
        ' Output the greatest increase, decrease, and volume in columns O, P, Q
        ws.Range("O2:O4").Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(4, 16).Value = volumeTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 17).Value = greatestVolume
        
    Next ws
End Sub

