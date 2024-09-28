Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim outputRow As Long
    Dim quarterStartRow As Long
    Dim currentQuarter As Integer
    Dim i As Long
 
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Set up headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
 
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
 
        ' Find last row of data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
 
        ' Initialize variables
        outputRow = 2
        totalVolume = 0
        quarterStartRow = 2
        currentQuarter = DatePart("q", ws.Cells(2, "B").Value)
 
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if we're still in the same ticker and quarter
            If ws.Cells(i, "A").Value = ws.Cells(i + 1, "A").Value And _
               DatePart("q", ws.Cells(i, "B").Value) = DatePart("q", ws.Cells(i + 1, "B").Value) Then
                totalVolume = totalVolume + ws.Cells(i, "G").Value
            Else
                ' Calculate quarterly change
                ticker = ws.Cells(i, "A").Value
                openingPrice = ws.Cells(quarterStartRow, "C").Value
                closingPrice = ws.Cells(i, "F").Value
                totalVolume = totalVolume + ws.Cells(i, "G").Value
 
                quarterlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = quarterlyChange / openingPrice
                Else
                    percentChange = 0
                End If
 
                ' Output results
                ws.Cells(outputRow, "I").Value = ticker
                ws.Cells(outputRow, "J").Value = quarterlyChange
                ws.Cells(outputRow, "K").Value = percentChange
                ws.Cells(outputRow, "L").Value = totalVolume
 
                ' Format cells
                ws.Cells(outputRow, "K").NumberFormat = "0.00%"
                If quarterlyChange >= 0 Then
                    ws.Cells(outputRow, "J").Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(outputRow, "J").Interior.Color = RGB(255, 0, 0)
                End If
 
                ' Reset for next ticker/quarter
                outputRow = outputRow + 1
                totalVolume = 0
                quarterStartRow = i + 1
            End If
        Next i
 
        ' Find greatest values
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim increaseTickerx As String
        Dim decreaseTicker As String
        Dim volumeTicker As String
 
        greatestIncrease = WorksheetFunction.Max(ws.Range("K:K"))
        greatestDecrease = WorksheetFunction.Min(ws.Range("K:K"))
        greatestVolume = WorksheetFunction.Max(ws.Range("L:L"))
 
        increaseTicker = ws.Cells(WorksheetFunction.Match(greatestIncrease, ws.Range("K:K"), 0), "I").Value
        decreaseTicker = ws.Cells(WorksheetFunction.Match(greatestDecrease, ws.Range("K:K"), 0), "I").Value
        volumeTicker = ws.Cells(WorksheetFunction.Match(greatestVolume, ws.Range("L:L"), 0), "I").Value
 
        ' Output greatest values
        ws.Range("P2").Value = increaseTicker
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
 
        ws.Range("P3").Value = decreaseTicker
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
 
        ws.Range("P4").Value = volumeTicker
        ws.Range("Q4").Value = greatestVolume
 
        ' Autofit columns
        ws.Columns("I:Q").AutoFit
    Next ws
End Sub
