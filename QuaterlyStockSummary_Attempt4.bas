Attribute VB_Name = "QuaterlyStockSummary_Attempt4"
Sub SummarizeStockData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1") ' Change to the sheet you're working on

    ' Find the last row with data in column A (ticker column)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Define output start row and set headers
    Dim outputRow As Long
    outputRow = 2 ' First output row after headers
    ws.Range("I1:L1").Value = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")
    
    ' Collection to hold unique tickers
    Dim uniqueTickers As Collection
    Set uniqueTickers = New Collection
    
    ' Collect unique tickers
    Dim i As Long, ticker As String
    On Error Resume Next
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        If ticker <> "" Then uniqueTickers.Add ticker, CStr(ticker)
    Next i
    On Error GoTo 0
    
    ' Loop through each unique ticker to process calculations
    Dim t As Variant, firstOpen As Double, lastClose As Double, totalVol As Double
    Dim quarterlyChange As Double, percentChange As Double
    
    For Each t In uniqueTickers
        totalVol = 0
        firstOpen = 0
        lastClose = 0
        Dim firstOpenFound As Boolean
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
        With ws.Range("I" & outputRow & ":L" & outputRow)
            ' Write Ticker
            .Cells(1, 1).Value = t
            
            ' Write Quarterly Change with number format and color formatting
            .Cells(1, 2).Value = quarterlyChange
            .Cells(1, 2).NumberFormat = "0.00"
            .Cells(1, 2).Interior.Color = IIf(quarterlyChange > 0, RGB(144, 238, 144), IIf(quarterlyChange < 0, RGB(255, 99, 71), xlNone))
            
            ' Write Percent Change with number format
            .Cells(1, 3).Value = percentChange
            .Cells(1, 3).NumberFormat = "0.00"
            
            ' Write Total Volume
            .Cells(1, 4).Value = totalVol
        End With
        
        ' Move to the next output row
        outputRow = outputRow + 1
    Next t
End Sub

