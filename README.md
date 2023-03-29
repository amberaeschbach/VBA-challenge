# VBA-challenge
homework 2 for bootcamp
Sub Main()
    
    Dim rCount As Long
    Dim currentTicker As String
    'Dim j As Integer
    Dim openValue As Double, closeValue As Double, totalStock As Double
    Dim greatInc As Double, greatDec As Double, greatStock As Double
    Dim greatIncTicker As String, greatDecTicker As String, greatStockTicker As String
    Dim FormatRange As Range


For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row
        
        totalStock = 0
        j = 2
        
        openValue = ws.Cells(2, 3).Value
    
        For k = 2 To RowCount
            totalStock = totalStock + ws.Cells(k, 7).Value
            If ws.Cells(k + 1, 1).Value <> ws.Cells(k, 1).Value Then
                closeValue = ws.Cells(k, 6).Value
    
                ws.Cells(j, 9).Value = ws.Cells(k, 1).Value
                ws.Cells(j, 10).Value = closeValue - openValue
                ws.Cells(j, 11).Value = (closeValue - openValue) / openValue
                ws.Cells(j, 10).NumberFormat = "0.00"
                ws.Cells(j, 11).NumberFormat = "0.00%"
                ws.Cells(j, 12).Value = totalStock
                
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                End If
                    
                If ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
                    
                j = j + 1
                totalStock = 0
                openValue = ws.Cells(k + 1, 3).Value
            End If
        Next k
        
        RowCount = Cells(Rows.Count, 9).End(xlUp).Row
        
        greatInc = 0
        greatDec = 0
        greatStock = 0
        
        For k = 2 To RowCount
            If ws.Cells(k, 11).Value > greatInc Then
                greatIncTicker = ws.Cells(k, 9).Value
                greatInc = ws.Cells(k, 11).Value
            End If
            If ws.Cells(k, 11).Value < greatDec Then
                greatDecTicker = ws.Cells(k, 9).Value
                greatDec = ws.Cells(k, 11).Value
            End If
            If ws.Cells(k, 12).Value > greatStock Then
                greatStockTicker = ws.Cells(k, 9).Value
                greatStock = ws.Cells(k, 12).Value
            End If
            
        Next k
        
        
        
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
            
        ws.Range("P2").Value = greatIncTicker
        ws.Range("P3").Value = greatDecTicker
        ws.Range("P4").Value = greatStockTicker
        ws.Range("Q2").Value = greatInc
        ws.Range("Q3").Value = greatDec
        ws.Range("Q4").Value = greatStock
        


Next ws

End Sub
