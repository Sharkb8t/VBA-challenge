Attribute VB_Name = "QuaterlyStockSummary_Attempt2"
Sub SummarizeStockData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1") ' Change to the sheet you're working on

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim outputRow As Long
    outputRow = 2 ' Starting row for output
    
    ' Add headers to output columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim ticker As String
    Dim firstOpen As Double, lastClose As Double
    Dim totalVol As Double, quarterlyChange As Double
    Dim percentChange As Double
    Dim firstRow As Long, lastRowTicker As Long
    Dim uniqueTickers As Collection
    Set uniqueTickers = New Collection
    
    Dim i As Long
    
    ' Find all unique tickers
    On Error Resume Next
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        If ticker <> "" Then
            uniqueTickers.Add ticker, CStr(ticker)
        End If
    Next i
    On Error GoTo 0
    
    ' Loop through each unique ticker and calculate the required values
    Dim t As Variant
    For Each t In uniqueTickers
        ' Get the first and last row for this ticker
        firstRow = 0
        lastRowTicker = 0
        totalVol = 0
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = t Then
                If firstRow = 0 Then
                    firstRow = i ' First appearance of this ticker
                End If
                lastRowTicker = i ' Update last appearance
                totalVol = totalVol + ws.Cells(i, 7).Value ' Summing volume
            End If
        Next i
        
        ' Calculate the first <open> and last <close>
        firstOpen = ws.Cells(firstRow, 3).Value
        lastClose = ws.Cells(lastRowTicker, 6).Value
        
        ' Calculate quarterly change and percent change
        quarterlyChange = lastClose - firstOpen
        percentChange = (quarterlyChange / firstOpen) * 100
        
        ' Write the results to the output columns
        ws.Cells(outputRow, 9).Value = t ' Ticker
        ws.Cells(outputRow, 10).Value = quarterlyChange ' Quarterly Change
        ws.Cells(outputRow, 11).Value = percentChange ' Percent Change
        ws.Cells(outputRow, 12).Value = totalVol ' Total Volume
        
        outputRow = outputRow + 1
    Next t
End Sub

