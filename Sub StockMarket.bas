Sub StockMarket()
'Loop through all worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    ' Set an initial variable for holding the ticker
    Dim TickerRow As String
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Define last row
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim Totalvolume As Double
    Totalvolume = 0
    
    ' Create new column and row titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % decrease"
    Range("O4").Value = "Greatest total volume"
    
    ws.Columns("K:K").NumberFormat = "0.00""%"""
    ws.Columns("Q:Q").NumberFormat = "0.00""%"""
    ws.Columns.AutoFit

    ' Define variables
    Dim openingprice As Double
    Dim closingprice As Double
    Dim Percentagechange As Double
    Dim maxPercentIncrease As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecrease As Double
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolume As Double
    Dim maxTotalVolumeTicker As String
    

    ' Loop through all tickers in rows
    For i = 2 To lastrow
        If openingprice = 0 Then
            openingprice = Cells(i, 3).Value
        End If
        
        ' Check if we are still within the same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Set the current Ticker
            TickerRow = Cells(i, 1).Value
            Totalvolume = Totalvolume + Cells(i, 7).Value

            closingprice = Cells(i, 6).Value
            Yearlychanged = closingprice - openingprice
            
            
            If openingprice <> 0 Then
                Percentagechange = (Yearlychanged / openingprice) * 100
            End If
            
            openingprice = 0
            
            ' Print the values in the Summary Table
            Range("I" & Summary_Table_Row).Value = TickerRow
            Range("J" & Summary_Table_Row).Value = Yearlychanged
            Range("K" & Summary_Table_Row).Value = Percentagechange
            Range("L" & Summary_Table_Row).Value = Totalvolume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Totalvolume = 0
        Else
            Totalvolume = Totalvolume + Cells(i, 7).Value
        End If
    Next i
    
     'Set cell color based on yearly change
     For i = 2 To lastrow
            If Cells(i, 10).Value <= 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
            ElseIf Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            End If
            Next i
             Cells(i, 10).Interior.ColorIndex = xlNone

    ' Find maximum percentage increase after the first loop
    maxPercentIncrease = 0 ' Reset to zero before searching
    For i = 2 To lastrow
        If (Cells(i, 11).Value) > (maxPercentIncrease) Then
            maxPercentIncrease = Cells(i, 11).Value
            maxPercentIncreaseTicker = Cells(i, 9).Value
        End If
    Next i
    
    maxPercentDecrease = 0 
    For i = 2 To lastrow
        If (Cells(i, 11).Value) < (maxPercentDecrease) Then
            maxPercentDecrease = Cells(i, 11).Value
            maxPercentDecreaseTicker = Cells(i, 9).Value
        End If
    Next i
    
    maxTotalVolume = 0 
    For i = 2 To lastrow
        If (Cells(i, 12).Value) > (maxTotalVolume) Then
            maxTotalVolume = Cells(i, 12).Value
            maxTotalVolumeTicker = Cells(i, 9).Value
        End If
    Next i

    ' Output maximum values
    Range("P2").Value = maxPercentIncreaseTicker
    Range("Q2").Value = maxPercentIncrease
    
    Range("P3").Value = maxPercentDecreaseTicker
    Range("Q3").Value = maxPercentDecrease
    
    Range("P4").Value = maxTotalVolumeTicker
    Range("Q4").Value = maxTotalVolume
    Next ws
End Sub