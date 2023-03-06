Attribute VB_Name = "Module1"
Sub Stocks()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
        Dim rowCount As Integer
        Dim lastrow As Long
        rowCount = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim total As Double
        Dim first As Double
        total = 0
        first = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        For i = 2 To lastrow
        
            total = total + ws.Cells(i, 7).Value
            
            If first = 0 Then
                first = ws.Cells(i, 3).Value
            End If
            
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ws.Cells(rowCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(rowCount, 12).Value = total
                
                ws.Cells(rowCount, 10).Value = ws.Cells(i, 6).Value - first
                ws.Cells(rowCount, 11).Value = FormatPercent(ws.Cells(rowCount, 10).Value / first)
                
                If (ws.Cells(rowCount, 10).Value > 0) Then
                    ws.Cells(rowCount, 10).Interior.ColorIndex = 4
                    ws.Cells(rowCount, 11).Interior.ColorIndex = 4
                ElseIf (ws.Cells(rowCount, 10).Value < 0) Then
                    ws.Cells(rowCount, 10).Interior.ColorIndex = 3
                    ws.Cells(rowCount, 11).Interior.ColorIndex = 3
                End If
                
                total = 0
                first = 0
                rowCount = rowCount + 1
            End If
        Next i
        
        
        Dim increase As Double
        Dim incTicker As String
        Dim decrease As Double
        Dim decTicker As String
        Dim largestTotal As Double
        Dim totalTicker As String
        
        increase = -100
        decrease = 100
        largestTotal = 0
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
            
        For j = 2 To rowCount
            If (ws.Cells(j, 11).Value > increase) Then
                increase = ws.Cells(j, 11).Value
                incTicker = ws.Cells(j, 9).Value
            End If
            
            If (ws.Cells(j, 11).Value < decrease) Then
                decrease = ws.Cells(j, 11).Value
                decTicker = ws.Cells(j, 9).Value
            End If
            
            If (ws.Cells(j, 12).Value > largestTotal) Then
                largestTotal = ws.Cells(j, 12).Value
                totalTicker = ws.Cells(j, 9).Value
            End If
        
        Next j
        
        ws.Range("P2").Value = incTicker
        ws.Range("Q2").Value = FormatPercent(increase)
        ws.Range("P3").Value = decTicker
        ws.Range("Q3").Value = FormatPercent(decrease)
        ws.Range("P4").Value = totalTicker
        ws.Range("Q4").Value = largestTotal
        
        
    Next ws

End Sub


