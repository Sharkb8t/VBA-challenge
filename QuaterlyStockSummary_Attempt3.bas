Attribute VB_Name = "QuaterlyStockSummary_Attempt3"
Sub SummarizeStockData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1") ' Sheet to work on

    ' Find the last row with data in column A (ticker column)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Define output start row and set headers
    Dim outputRow As Long
    outputRow = 2 ' First output row after headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Collection to hold unique tickers
    Dim uniqueTickers As Collection
    Set uniqueTickers = New Collection
    
    ' Loop through the tickers and collect unique values
    Dim i As Long, ticker As String
    On Error Resume Next
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        If ticker <> "" Then uniqueTickers.Add ticker, CStr(ticker)
    Next i
    On Error GoTo 0
    
    ' Loop through each unique ticker to process calculations
    Dim t As Variant, firstOpen As Double, lastClose As Double, totalVol As Double
    Dim quarterlyChange As Double, percentChange As Double, firstRow As Long, lastRowTicker As Long
    
    For Each t In uniqueTickers
        firstRow = 0
        lastRowTicker = 0
        totalVol = 0
        
        ' Loop through all rows to find first and last occurrence of ticker and sum volume
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = t Then
                If firstRow = 0 Then firstRow = i ' First appearance
                lastRowTicker = i ' Track last appearance
                totalVol = totalVol + ws.Cells(i, 7).Value ' Sum volume
            End If
        Next i
        
        ' Calculate values
        firstOpen = ws.Cells(firstRow, 3).Value
        lastClose = ws.Cells(lastRowTicker, 6).Value
        quarterlyChange = lastClose - firstOpen
        percentChange = Round((quarterlyChange / firstOpen) * 100, 2) ' Percent change rounded to 2 decimals
        
        ' Write to output and apply formatting in one step
        With ws
            ' Write Ticker
            .Cells(outputRow, 9).Value = t
            
            ' Write Quarterly Change and apply conditional formatting for color
            .Cells(outputRow, 10).Value = quarterlyChange
            Select Case quarterlyChange
                Case Is > 0
                    .Cells(outputRow, 10).Interior.Color = RGB(144, 238, 144) ' Green
                Case Is < 0
                    .Cells(outputRow, 10).Interior.Color = RGB(255, 99, 71) ' Red
                Case Else
                    .Cells(outputRow, 10).Interior.ColorIndex = xlNone ' No fill
            End Select
            
            ' Write Percent Change and ensure 2 decimal places
            .Cells(outputRow, 11).Value = percentChange
            .Cells(outputRow, 11).NumberFormat = "0.00"
            
            ' Write Total Volume
            .Cells(outputRow, 12).Value = totalVol
        End With
        
        ' Increment output row
        outputRow = outputRow + 1
    Next t
End Sub


