Sub VBAChallenge():

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets

        Dim lngLastRowData As Long
        Dim i As Long
        Dim strCurrentTicker As String
        Dim lngSummaryTickerCount As Long
        Dim dblOpenPriceFirstDay As Double
        Dim dblOpenPrice As Double
        Dim dblClosePrice As Double
        Dim dblYearlyChange As Double
        Dim dblPercentChange As Double
        Dim dblTotalYearlyVolume As Double
    
        dblOpenPriceFirstDay = ws.Range("C" & 2).Value 'Setting up the Open price of first day of first ticker
        dblTotalYearlyVolume = 0 ' Setting total stock volume to 0
        lngSummaryTickerCount = 2
        lngLastRowData = ws.Range("A" & Rows.Count).End(xlUp).Row
        
        For i = 2 To lngLastRowData
                        
            If ws.Range("A" & i).Value <> ws.Range("A" & i + 1).Value Then
                
                strCurrentTicker = ws.Range("A" & i).Value
                dblClosePrice = ws.Range("F" & i).Value
                dblYearlyChange = dblClosePrice - dblOpenPriceFirstDay
                dblPercentChange = (dblYearlyChange / dblOpenPriceFirstDay)
                dblTotalYearlyVolume = dblTotalYearlyVolume + ws.Range("G" & i)
                ws.Range("I" & lngSummaryTickerCount).Value = strCurrentTicker
                ws.Range("J" & lngSummaryTickerCount).Value = dblYearlyChange
                ws.Range("K" & lngSummaryTickerCount).Value = dblPercentChange
                ws.Range("L" & lngSummaryTickerCount).Value = dblTotalYearlyVolume
                dblOpenPriceFirstDay = ws.Range("C" & i + 1).Value 'Setting the open price of first day for next ticker
                dblTotalYearlyVolume = 0 ' Resetting the total stock volume to 0 for next ticker
                lngSummaryTickerCount = lngSummaryTickerCount + 1
                
            Else
             
                dblTotalYearlyVolume = dblTotalYearlyVolume + ws.Range("G" & i)
                
            End If
            
        Next i
        'Creating columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Columns("K:K").NumberFormat = "0.00%"
        
        
        Dim lngLastRowSummary As Long
        
        lngLastRowSummary = ws.Range("I" & Rows.Count).End(xlUp).Row
        
        'Applying conditional formatting for Columns Yearly Change and % Change
        
        For i = 2 To lngLastRowSummary
        
            For j = 10 To 11
        
                If ws.Cells(i, j).Value >= 0 Then
                
                    ws.Cells(i, j).Interior.Color = RGB(0, 255, 0)
                
                Else
            
                    ws.Cells(i, j).Interior.Color = RGB(255, 0, 0)
                
                End If
                
            Next j
                
        Next i
        
        Dim dblMaxIncrease As Double
        Dim lngRowMaxIncrease As Long
        Dim dblMaxDecrease As Double
        Dim lngRowMaxDecrease As Long
        Dim dblMaxVolume As Double
        Dim lngRowMaxVolume As Long
        Dim strTickerMaxIncrease As String
        Dim strTickerMaxDecrease As String
        Dim strTickerMaxVolume As String
                
        'Calculating the max and min value of %change column and max value of Total Stock Volume column. I have called the
        'the functions GetMaxRow and GetMinRow which I created to get the respective row numbers
       
        dblMaxIncrease = WorksheetFunction.Max(ws.Range("K2:K" & lngLastRowSummary))
        lngRowMaxIncrease = GetMaxRow(ws.Range("K2:K" & lngLastRowSummary))
        
        dblMaxDecrease = WorksheetFunction.Min(ws.Range("K2:K" & lngLastRowSummary))
        lngRowMaxDecrease = GetMinRow(ws.Range("K2:K" & lngLastRowSummary))
        
        dblMaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & lngLastRowSummary))
        lngRowMaxVolume = GetMaxRow(ws.Range("L2:L" & lngLastRowSummary))
        
        'Finding the tickers for respective max and min values
        strTickerMaxIncrease = ws.Range("I" & lngRowMaxIncrease).Value
        strTickerMaxDecrease = ws.Range("I" & lngRowMaxDecrease).Value
        strTickerMaxVolume = ws.Range("I" & lngRowMaxVolume).Value
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P2").Value = strTickerMaxIncrease
        ws.Range("P3").Value = strTickerMaxDecrease
        ws.Range("P4").Value = strTickerMaxVolume
        ws.Range("Q2").Value = dblMaxIncrease
        ws.Range("Q3").Value = dblMaxDecrease
        ws.Range("Q4").Value = dblMaxVolume
        ws.Columns("I:Q").AutoFit
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
       
        
    Next ws
        
End Sub

Function GetMaxRow(myRange As Range) As Long

    Dim rng As Range
    Dim maxVal As Double
    Dim maxRow As Long

    maxVal = WorksheetFunction.Min(myRange)
    maxRow = 0

    For Each rng In myRange

        If rng.Value > maxVal Then

            maxRow = rng.Row
            maxVal = rng.Value

        End If

    Next rng

    GetMaxRow = maxRow

End Function

Function GetMinRow(myRange As Range) As Long

    Dim rng As Range
    Dim minVal As Double
    Dim minRow As Long

    minVal = WorksheetFunction.Max(myRange)
    minRow = 0

    For Each rng In myRange

        If rng.Value < minVal Then

            minRow = rng.Row
            minVal = rng.Value

        End If

    Next rng

    GetMinRow = minRow

End Function

