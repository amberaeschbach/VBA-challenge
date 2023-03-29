# VBA-challenge
        
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
                    
                j = j + 1
                totalStock = 0
                openValue = ws.Cells(k + 1, 3).Value
            End If
        Next k
 FOR THE ABOVE CODE, I WORKED WITH A TUTOR TO EDIT THE CODE I HAD AND THIS IS WHAT WE CAME UP WITH 
 
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
 FOR THE ABOVE CODE, I WORKED WITH A FRIEND OF MINE TO EDIT MY CODE, AND THIS IS WHAT WE CAME UP WITH 
