Attribute VB_Name = "Module1"
Sub stockAnalysis():

    Dim total As Double
    Dim row As Long
    Dim rowCount As Long
    Dim change As Double
    Dim yearlyChange As Double
    Dim summaryTableRow As Long
    Dim stockStartRow As Long
    
    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        summaryTableRow = 0
        total = 0
        yearlyChange = 0
        stockStartRow = 2
        
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        For row = 2 To rowCount
        
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                total = total + ws.Cells(row, 7).Value
                If total = 0 Then
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                Else
                    If ws.Cells(stockStartRow, 3).Value = 0 Then
                        For findValue = stockStartRow To row
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                stockStartRow = findValue
                                Exit For
                            End If
                        Next findValue
                    End If
                    
                    yearlyChange = (Cells(row, 6).Value - ws.Cells(stockStartRow, 3).Value)
                    percentChange = yearlyChange / Cells(stockStartRow, 3).Value
                    
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
                    ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###"
                    
                    If yearlyChange > 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                    ElseIf yearlyChange < 0 Then
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    End If
                
                
                End If
                
                total = 0
                yearlyChange = 0
                summaryTableRow = summaryTableRow + 1
                 
            Else
                total = total + ws.Cells(row, 7).Value
            End If
    
        Next row
    
        ws.Columns("A:Q").AutoFit
    Next ws
End Sub
