Sub Test()

 For Each ws In Worksheets
  
  ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
  Dim WorksheetName As String

  ws.Activate
  'Get the last row number
  Dim lastRow As Long
  lastRow = 2
  
  While Trim(Cells(lastRow, 1).Value) <> ""
    lastRow = lastRow + 1
  Wend

Dim isFirstOne As Boolean
    isFirstOne = True
       
    Dim openValue As Double
    Dim closeValue As Double
    Dim currentTicker As String
    Dim totalchange As Double
    Dim pctChange As Double
    Dim totalVolume As Double
    
    openValue = 0
    closeValue = 0
    currentTicker = "A"
    totalchange = 0
    pctChange = 0
    totalVolume = 0
        
    Dim currentTickerIndex As Long
    currentTickerIndex = 2
        
    For i = 2 To lastRow
        currentTicker = Cells(i, 1).Value
        
        isLastOne = currentTicker <> Cells(i + 1, 1).Value
        
        totalVolume = totalVolume + Cells(i, 7).Value
        
        If isFirstOne Then
            openValue = Cells(i, 3).Value
            isFirstOne = False
        End If
            
        If isLastOne Then
            closeValue = Cells(i, 3).Value
            totalchange = closeValue - openValue
            
            If openValue = 0 Then
                pctChange = 0
            Else
                pctChange = totalchange / openValue
            End If
            
            'set the total change, pct change, and total volume numbers in the excel sheet
            'set these values at row currentTickerIndex
            'so basically, do this:
            Cells(1, 9) = "Ticker"
            Cells(1, 10) = "Yearly Change"
            Cells(1, 11) = "Percent Change"
            Cells(1, 12) = "Total Stock Volume"
           
            
            Range("I" & currentTickerIndex).Value = currentTicker
            Range("J" & currentTickerIndex).Value = totalchange
            
            If totalchange > 0 Then
                Range("J" & currentTickerIndex).Interior.Color = vbGreen
            ElseIf totalchange < 0 Then
                Range("J" & currentTickerIndex).Interior.Color = vbRed
            End If
            
            Range("K" & currentTickerIndex).Value = pctChange
            Range("K" & currentTickerIndex).NumberFormat = "0.00%"
            Range("L" & currentTickerIndex).Value = totalVolume
            
            currentTickerIndex = currentTickerIndex + 1
            totalVolume = 0
            isLastOne = False
            isFirstOne = True
        End If
    Next i
    
    greatestPctChange = Cells(2, 11).Value
    greatestPctChangeTicker = "A"
    
    smallestPctChange = greatestPctChange
    smallestPctChangeTicker = "A"
    
    greatestTotalVolume = Cells(2, 12).Value
    greatestTotalVolumeTicker = "A"
    
    For i = 2 To currentTickerIndex + 1
        If Cells(i, 11).Value > greatestPctChange Then
            greatestPctChange = Cells(i, 11).Value
            greatestPctChangeTicker = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value < smallestPctChange Then
            smallestPctChange = Cells(i, 11).Value
            smallestPctChangeTicker = Cells(i, 9).Value
        End If
            
        If Cells(i, 12).Value > greatestTotalVolume Then
            greatestTotalVolume = Cells(i, 12).Value
            greatestTotalVolumeTicker = Cells(i, 9).Value
        End If
                
        Range("N" & 2).Value = "Greatest % Increase"
        Range("O" & 2).Value = greatestPctChangeTicker
        Range("P" & 2).Value = greatestPctChange
        Range("P" & 2).NumberFormat = "0.00%"
        
        Range("N" & 3).Value = "Greatest % Decrease"
        Range("O" & 3).Value = smallestPctChangeTicker
        Range("P" & 3).Value = smallestPctChange
        Range("P" & 3).NumberFormat = "0.00%"
        
        Range("N" & 4).Value = "Greatest Total Volume"
        Range("O" & 4).Value = greatestTotalVolumeTicker
        Range("P" & 4).Value = greatestTotalVolume
    Next i
  
  Next ws

End Sub


