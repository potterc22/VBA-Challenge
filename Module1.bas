Attribute VB_Name = "Module1"
Sub VBAofWallStreet()
    For Each ws In Worksheets
        
    'Create headings for data summary
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    'Create headings for hard solution
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        
    'Declare Variables
        Dim ticker As String
        Dim YearlyChange As Double
        Dim Count As Long
        Dim StockVolume As Double
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim PercentChange As Double
    'Declare Variables for Hard solution
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotal As Double
        
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotal = 0
        StockVolume = 0
        'Use count when printing values
        Count = 2
        
        
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'loop through each row of each sheet
        For i = 2 To Lastrow
        'loop through each stock adding up their total volume
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            'Form the summary table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                'Print ticker symbol
                ws.Range("I" & Count).Value = ticker
                
                'Print Total Stock Volume
                ws.Range("L" & Count).Value = StockVolume
                'Reset StockVolume for next stock
                StockVolume = 0
                
                'Calculate Yearly Change
                YearOpen = ws.Range("C" & i)
                YearClose = ws.Range("F" & i)
                YearlyChange = YearClose - YearOpen
                'Print Yearly Change
                ws.Range("J" & Count).Value = YearlyChange
    
                'Calculate Percent Change
                If YearOpen = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / YearOpen
                End If
                'Print Percent Change
                ws.Range("K" & Count).Value = PercentChange
                'Format the cells to be percentages
                ws.Range("K" & Count).Style = "Percent"
                
                'Conditional formatting: Positive = Green, Negative = Red, No Change = Yellow
                If YearlyChange > 0 Then
                    ws.Range("J" & Count).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & Count).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Count).Interior.ColorIndex = 6
                End If
            Count = Count + 1
            End If
        Next i
        'Autofit the table columns
        ws.Columns("I:P").AutoFit
        
        'Hard solution
        'Find the last row of the summary table
        LastSummaryRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Calculate Greatest Values
        For j = 2 To LastSummaryRow
            'Calculate Greatest % Increase
            If ws.Range("K" & j).Value > GreatestIncrease Then
                GreatestIncrease = ws.Range("K" & j).Value
                'Print Greatest % Increase
                ws.Range("P2").Value = GreatestIncrease
                'Print associated ticker
                ws.Range("O2").Value = ws.Range("I" & j).Value
            End If
            'Calculate Greatest % Decrease
            If ws.Range("K" & j).Value < GreatestDecrease Then
                GreatestDecrease = ws.Range("K" & j).Value
                'Print Greatest % Decrease
                ws.Range("P3").Value = GreatestDecrease
                'Print associated ticker
                ws.Range("O3").Value = ws.Range("I" & j).Value
            End If
            'Calculate Greatest total volume
            If ws.Range("L" & j).Value > GreatestTotal Then
                GreatestTotal = ws.Range("L" & j).Value
                'Print Greatest total Volume
                ws.Range("P4").Value = GreatestTotal
                'Print associated ticker
                ws.Range("O4").Value = ws.Range("I" & j).Value
            End If
        Next j
        ws.Range("P2:P3").Style = "Percent"
    Next ws
End Sub
