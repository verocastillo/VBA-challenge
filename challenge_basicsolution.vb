'  Homework #2: VBA Challenge
'
' This code condenses the information from real stock data
' in order to analyze it. This script contains the basic solution.
'
Sub EasyChallenge()

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
        Dim PercentC As Variant
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
                ws.Cells(NewRow, 11).Style = "Percent"
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
    Next ws

End Sub
