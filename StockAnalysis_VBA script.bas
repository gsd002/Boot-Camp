Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Integer
    
    For Each ws In ThisWorkbook.Worksheets
    
        ' Initialize summary row
        SummaryRow = 2

        ' Set headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        ' Determine the last row of data
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Loop through all rows of data
        For i = 2 To LastRow
        
            ' Check if the current ticker is different from the previous ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosingPrice = ws.Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
                ' Output the results to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Conditional Formatting for YearlyChange
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Conditional Formatting for PercentChange
                If PercentChange > 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0)
                ElseIf PercentChange < 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0)
                End If

                ' Move to the next summary row
                SummaryRow = SummaryRow + 1
            
                ' Reset variables for the next ticker
                OpeningPrice = 0
                TotalVolume = 0

            Else
                ' If it's the same ticker, accumulate the volume
                If OpeningPrice = 0 Then
                    OpeningPrice = ws.Cells(i, 3).Value
                End If
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i

        ' Formatting
        ws.Range("J2:J" & SummaryRow - 1).NumberFormat = "0.00"
        ws.Range("K2:K" & SummaryRow - 1).NumberFormat = "0.00%"
        Next ws
    
End Sub

Sub FindGreatestValues()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim rngPercentChange As Range, rngTotalVolume As Range
    Dim GreatestInc As Double, GreatestDec As Double, GreatestVol As Double
    Dim GreatestIncCell As Range, GreatestDecCell As Range, GreatestVolCell As Range
    Dim GreatestIncTicker As String, GreatestDecTicker As String, GreatestVolTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
    
        ' Determine the last row of data in the worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set Ranges for "Percent Change" and "Total Volume"
        Set rngPercentChange = ws.Range("K2:K" & LastRow)
        Set rngTotalVolume = ws.Range("L2:L" & LastRow)

        ' Find the maximum and minimum values for percent change
        GreatestInc = Application.WorksheetFunction.Max(rngPercentChange)
        GreatestDec = Application.WorksheetFunction.Min(rngPercentChange)
    
        ' Find the maximum value for total volume
        GreatestVol = Application.WorksheetFunction.Max(rngTotalVolume)
        
        ' Identify the cells associated with these values
        Set GreatestIncCell = rngPercentChange.Find(GreatestInc)
        Set GreatestDecCell = rngPercentChange.Find(GreatestDec)
        Set GreatestVolCell = rngTotalVolume.Find(GreatestVol)

        ' Check if cells are found and then extract tickers
        GreatestVolTicker = GreatestVolCell.Offset(0, -3).Value


        ' Output results in a table starting from cell O1
        ws.Range("O1").Value = "Criteria"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = GreatestIncTicker
        ws.Range("Q2").Value = GreatestInc
    
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = GreatestDecTicker
        ws.Range("Q3").Value = GreatestDec
    
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = GreatestVolTicker
        ws.Range("Q4").Value = GreatestVol
        
        ' Formatting
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    
    Next ws

End Sub


