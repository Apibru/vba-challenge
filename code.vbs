Sub StockDataAnalysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim SummaryTableRow As Integer
    
    ' Variables to store greatest values and corresponding tickers
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim TickerGreatestPercentIncrease As String
    Dim TickerGreatestPercentDecrease As String
    Dim TickerGreatestTotalVolume As String

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
        ' Set the initial value for the summary table row
        SummaryTableRow = 2
        
        ' Find the last row with data in the worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Create headers for the summary table
        ws.Cells(1, 9).Value = "<ticker>"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Cells(1, 11).Value = "Yearly Change ($)"
        ws.Cells(1, 12).Value = "Percent Change"
        
        ' Initialize values for greatest measures
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0
        
        ' Loop through rows to retrieve data
        For i = 2 To LastRow
        
            ' Check if we are at the last entry for the current ticker or if the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                ' Accumulate the total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                ' Set the closing price
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change
                YearlyChange = ClosePrice - OpenPrice
                
                ' Calculate the percent change
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = (YearlyChange / OpenPrice) * 100
                End If
                
                ' Output the data to the summary table
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                ws.Cells(SummaryTableRow, 10).Value = TotalVolume
                ws.Cells(SummaryTableRow, 11).Value = YearlyChange
                ws.Cells(SummaryTableRow, 12).Value = PercentChange
                
                ' Check for greatest measures
                If PercentChange > GreatestPercentIncrease Then
                    GreatestPercentIncrease = PercentChange
                    TickerGreatestPercentIncrease = Ticker
                ElseIf PercentChange < GreatestPercentDecrease Then
                    GreatestPercentDecrease = PercentChange
                    TickerGreatestPercentDecrease = Ticker
                End If
                
                If TotalVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalVolume
                    TickerGreatestTotalVolume = Ticker
                End If
                
                ' Move to the next row in the summary table
                SummaryTableRow = SummaryTableRow + 1
                
                ' Reset the total volume
                TotalVolume = 0
                
            Else
                ' If we are not at the last entry for the current ticker, accumulate the total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
            
            ' Check if this is the first entry for the current ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the opening price
                OpenPrice = ws.Cells(i, 3).Value
            End If
            
        Next i
        
        ' Apply conditional formatting for yearly change
        With ws.Range("K2:K" & SummaryTableRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Font
                .Color = -16752384 ' Green color
                .TintAndShade = 0
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13561798 ' Light green color
                .TintAndShade = 0
            End With
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Font
                .Color = -16383844 ' Red color
                .TintAndShade = 0
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615 ' Light red color
                .TintAndShade = 0
            End With
        End With
        
        ' Apply conditional formatting for percent change
        With ws.Range("L2:L" & SummaryTableRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Font
                .Color = -16752384 ' Green color
                .TintAndShade = 0
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13561798 ' Light green color
                .TintAndShade = 0
            End With
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Font
                .Color = -16383844 ' Red color
                .TintAndShade = 0
            End With
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615 ' Light red color
                .TintAndShade = 0
            End With
        End With
        
        ' Display the greatest values in the summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = TickerGreatestPercentIncrease
        ws.Cells(3, 16).Value = TickerGreatestPercentDecrease
        ws.Cells(4, 16).Value = TickerGreatestTotalVolume
        ws.Cells(2, 17).Value = GreatestPercentIncrease
        ws.Cells(3, 17).Value = GreatestPercentDecrease
        ws.Cells(4, 17).Value = GreatestTotalVolume
        
    Next ws

End Sub