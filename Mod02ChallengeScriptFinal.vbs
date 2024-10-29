Attribute VB_Name = "Module1"
Sub stockAnalysis()

    Dim total As Double                  ' total stock volume
    Dim row As Long                      ' loop control variable for rows in a sheet
    Dim rowCount As Long                 ' holds the number of rows in a sheet
    Dim quarterlyChange As Double        ' holds the quarterly change for each stock
    Dim percentChange As Double          ' holds the percent change for each stock
    Dim summaryTableRow As Long          ' holds the rows of the summary table row
    Dim stockStartRow As Long            ' holds the start row of the stock's rows
    Dim startValue As Long               ' start row for a stock (location of first open)
    Dim lastTicker As String             ' finds the last ticker in the sheet
    Dim findValue As Long                ' for finding the first non-zero open value
    Dim ws As Worksheet                  ' worksheet variable for loop
    Dim e As Long                        ' loop variable for clearing extra data
    Dim Column As Long                   ' column variable in nested loop

    ' loop through all worksheets in the Excel workbook
    For Each ws In Worksheets
        
        ' Set the Title Row of the Summary Section
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Set up the title row of the Aggregate Section
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' initialize values
        summaryTableRow = 0
        total = 0
        quarterlyChange = 0
        stockStartRow = 2
        startValue = 2
        
        ' get the last row with data in column A
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        ' find the last ticker
        lastTicker = ws.Cells(rowCount, 1).Value
        
        ' loop through each row
        For row = 2 To rowCount
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                
                ' add to the total stock volume
                total = total + ws.Cells(row, 7).Value
                
                ' check if the total stock volume is 0
                If total = 0 Then
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    ws.Range("K" & 2 + summaryTableRow).Value = 0
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                Else
                    If ws.Cells(startValue, 3).Value = 0 Then
                        For findValue = startValue To row
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                startValue = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                    
                    quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
                    percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
                    
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    
                    ' color the Quarterly change based on its value
                    If quarterlyChange > 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                    ElseIf quarterlyChange < 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    End If
                End If
                
                ' reset values for the next ticker
                total = 0
                quarterlyChange = 0
                startValue = row + 1
                summaryTableRow = summaryTableRow + 1
                
            Else
                total = total + ws.Cells(row, 7).Value
            End If
        Next row

        ' Update the summary table row
        summaryTableRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
        
        ' find and clear extra data
        Dim lastExtraRow As Long
        lastExtraRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).row
        
        For e = summaryTableRow + 1 To lastExtraRow
            For Column = 9 To 12
                ws.Cells(e, Column).Value = ""
                ws.Cells(e, Column).Interior.ColorIndex = 0
            Next Column
        Next e
        
        ' Calculate and print aggregates
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 1))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 1))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 1))
        
        Dim greatestIncreaseRow As Long
        Dim greatestDecreaseRow As Long
        Dim greatestTotVolRow As Long
        
        greatestIncreaseRow = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & summaryTableRow + 1), 0) + 1
        greatestDecreaseRow = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & summaryTableRow + 1), 0) + 1
        greatestTotVolRow = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & summaryTableRow + 1), 0) + 1
        
        ws.Range("P2").Value = ws.Cells(greatestIncreaseRow, 9).Value
        ws.Range("P3").Value = ws.Cells(greatestDecreaseRow, 9).Value
        ws.Range("P4").Value = ws.Cells(greatestTotVolRow, 9).Value
        
        ' Format columns
        For s = 0 To summaryTableRow - 1
            ws.Range("J" & 2 + s).NumberFormat = "0.00"
            ws.Range("K" & 2 + s).NumberFormat = "0.00%"
            ws.Range("L" & 2 + s).NumberFormat = "#,###"
        Next s
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,###"
        
        ' Autofit
        ws.Columns("A:Q").AutoFit
    Next ws
    
End Sub

