Sub StockAnalysis()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTable As Range
    Dim outputRow As Long
    Dim maximumIncrease As Double
    Dim maximumDecrease As Double
    Dim maximumVolume As Double
    Dim maximumIncreaseTicker As String
    Dim maximumDecreaseTicker As String
    Dim maximimVolumeTicker As String
    
    Set ws = ThisWorkbook.Sheets("2018")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    Set summaryTable = ws.Range("I1:L1")
    
    outputRow = 2
    
    maximumIncrease = 0
    maximumDecrease = 0
    maximumVolume = 0
    maximumIncreaseTicker = ""
    maximumDecreaseTicker = ""
    maximimVolumeTicker = ""
    
    For i = 2 To lastRow
        
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            totalVolume = 0
        End If
        
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            closePrice = ws.Cells(i, 6).Value
            
            yearlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = yearlyChange / openPrice
            Else
                percentChange = 0
            End If
            
            summaryTable.Cells(outputRow, 1).Value = ticker
            summaryTable.Cells(outputRow, 2).Value = yearlyChange
            summaryTable.Cells(outputRow, 3).Value = percentChange
            summaryTable.Cells(outputRow, 4).Value = totalVolume
            
            
            Select Case summaryTable.Cells(outputRow, 2).Value
                Case Is < 0
                    summaryTable.Cells(outputRow, 2).Interior.ColorIndex = 3
                Case Is > 0
                    summaryTable.Cells(outputRow, 2).Interior.ColorIndex = 4
                Case Else
                    summaryTable.Cells(outputRow, 2).Interior.ColorIndex = -4142
            End Select
            
            
            summaryTable.Cells(outputRow, 3).NumberFormat = "0.00%"
            
            
            If percentChange > maximumIncrease Then
                maximumIncrease = percentChange
                maximumIncreaseTicker = ticker
            End If
            If percentChange < maximumDecrease Then
                maximumDecrease = percentChange
                maximumDecreaseTicker = ticker
            End If
            If totalVolume > maximumVolume Then
                maximumVolume = totalVolume
                maximumVolumeTicker = ticker
            End If
            
            outputRow = outputRow + 1
            
            totalVolume = 0
        End If
    Next i
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = maximumIncreaseTicker
    ws.Cells(2, 17).Value = maximumIncrease
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = maximumDecreaseTicker
    ws.Cells(3, 17).Value = maximumDecrease
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = maximumVolumeTicker
    ws.Cells(4, 17).Value = maximumVolume
    
End Sub