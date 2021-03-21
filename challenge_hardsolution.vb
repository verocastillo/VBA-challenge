'  Homework #2: VBA Challenge plus Bonus
'
' This code condenses the information from real stock data
' in order to analyze it. This script contains the hard solution.
'
Sub HardChallenge()

     ' Do everything in every worksheet
    For Each ws In Worksheets

        ' Write headers for new table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Format new table
        For j = 9 To 12
            ws.Cells(1, j).Font.Bold = True
            ws.Cells(1, j).ColumnWidth = 16
        Next j
        
        ' Find last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Define variables
        Dim NewRow As Integer
        NewRow = 2
        Dim Ticker As String
        NewTicker = False
        Dim YearCh As Variant
        YearCh = 0
        Dim OpenC As Variant
        OpenC = ws.Cells(2, 3).Value
        Dim CloseC As Variant
        Dim PercentC As Double
        Dim TotalV As LongLong
        TotalV = 0

        ' Loop to fill information
        For i = 2 To LastRow
            ' Calculate total stock volume
            TotalV = TotalV + ws.Cells(i, 7).Value
            ' Check if the value changes between rows
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Put ticker value in the new table
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(NewRow, 9) = Ticker
                ' Operations to determine yearly change
                CloseC = ws.Cells(i, 6).Value
                YearCh = CloseC - OpenC
                ' Operations to determine percent change
                If OpenC <> 0 Then
                    PercentC = YearCh / OpenC
                Else
                    PercentC = 0
                End If
                ' Put yearly change value in the new table
                    If YearCh < 0 Then
                        ws.Cells(NewRow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(NewRow, 10).Interior.ColorIndex = 4
                    End If
                ws.Cells(NewRow, 10) = YearCh
                ' Put percent change value in new table
                ws.Cells(NewRow, 11) = PercentC
                ws.Cells(NewRow, 11).NumberFormat = "0.00%"
                ' Put total value in new table
                ws.Cells(NewRow, 12) = TotalV
                ' Change row
                NewRow = NewRow + 1
                ' Reset and define variables
                YearCh = 0
                OpenC = ws.Cells(i + 1, 3).Value
                TotalV = 0
            End If
        Next i
        
        ' Bonus exercises for the hard version
        ' Write headers and titles for new table
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ' Format new table
        For k = 2 To 4
            ws.Cells(k, 14).Font.Bold = True
            ws.Cells(k, 14).ColumnWidth = 15
        Next k
        For m = 15 To 16
            ws.Cells(1, m).Font.Bold = True
            ws.Cells(1, m).ColumnWidth = 15
        Next m
        
        ' Define variables
        Dim GInc As Double
        Dim GDec As Double
        Dim GTotal As LongLong
        Dim BTickerI As String
        Dim BTickerD As String
        Dim BTickerT As String
        
        ' Find last row
        BLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Define starting values
        GInc = ws.Cells(2, 11).Value
        GDec = ws.Cells(2, 11).Value
        GTotal = ws.Cells(2, 12).Value
        BTickerI = ws.Cells(2, 9).Value
        BTickerD = ws.Cells(2, 9).Value
        BTickerT = ws.Cells(2, 9).Value
        
        ' For loop to find table values
        For n = 2 To BLastRow
            'Find greatest change
            If GInc < ws.Cells(n, 11).Value Then
                GInc = ws.Cells(n, 11).Value
                BTickerI = ws.Cells(n, 9).Value
            End If
            'Find smallest change
            If GDec > ws.Cells(n, 11).Value Then
                GDec = ws.Cells(n, 11).Value
                BTickerD = ws.Cells(n, 9).Value
            End If
            'Find greatest value
            If GTotal < ws.Cells(n, 12).Value Then
                GTotal = ws.Cells(n, 12).Value
                BTickerT = ws.Cells(n, 9).Value
            End If
        Next n
        
        ' Put values in table
        ws.Cells(2, 15) = BTickerI
        ws.Cells(2, 16) = GInc
        ws.Cells(3, 15) = BTickerD
        ws.Cells(3, 16) = GDec
        ws.Cells(4, 15) = BTickerT
        ws.Cells(4, 16) = GTotal
        ws.Range("P2,P3").NumberFormat = "0.00%"
        
    Next ws

End Sub

