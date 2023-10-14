Sub StockAnalysis():

    Dim totalVol As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim lastRow As Long
    Dim i As Long
    Dim nextLine As Integer
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    nextLine = 2
    openPrice = 0
    closePrice = 0
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    lastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closePrice = Cells(i, 6).Value
            Cells(nextLine, 9).Value = Cells(i, 1).Value
            Cells(nextLine, 10).Value = closePrice - openPrice
            Cells(nextLine, 11).Value = ((closePrice - openPrice) / openPrice)
            Cells(nextLine, 11).NumberFormat = "0.00%"
            Cells(nextLine, 12).Value = totalVol + Cells(i, 7)
            nextLine = nextLine + 1
            openPrice = 0
            closePrice = 0
            totalVol = 0
            
        ElseIf openPrice = 0 Then
            openPrice = Cells(i, 3).Value
            totalVol = totalVol + Cells(i, 7)
        Else
            totalVol = totalVol + Cells(i, 7)
            
        End If
        
    Next i
    

    For i = 2 To lastRow
        If IsEmpty(Cells(i, 10).Value) Then
            Cells(i, 10).Interior.ColorIndex = 0
        ElseIf Cells(i, 10).Value > "0" Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < "0" Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    Next i
    
    
    For i = 2 To lastRow
        If Cells(i, 11).Value > greatestIncrease Then
            greatestIncrease = Cells(i, 11).Value
            Cells(2, 15).Value = Cells(i, 9).Value
            Cells(2, 16).Value = Cells(i, 11).Value
            Cells(2, 16).NumberFormat = "0.00%"
        End If
        
    Next i
    
    
        For i = 2 To lastRow
        If Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = Cells(i, 11).Value
            Cells(3, 15).Value = Cells(i, 9).Value
            Cells(3, 16).Value = Cells(i, 11).Value
            Cells(3, 16).NumberFormat = "0.00%"
        End If
        
    Next i
    
    
        For i = 2 To lastRow
        If Cells(i, 12).Value > greatestVolume Then
            greatestVolume = Cells(i, 12).Value
            Cells(4, 15).Value = Cells(i, 9).Value
            Cells(4, 16).Value = Cells(i, 12).Value
        End If
        
    Next i

End Sub
