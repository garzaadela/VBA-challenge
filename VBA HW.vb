Sub tickerloop()
    'Worksheet Object variable
    Dim CurrentWs As Worksheet
    Dim Results_Sheet As Boolean
    Need_Summary_Table_Header = True

    
    For Each CurrentWs In Worksheets
    

        'Variable for Ticker column in worksheet
        Dim ticker As String
        ticker = " "
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Percent_Change As Double
        Percent_Change = 0
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim summary_ticker_row As Long
        summary_ticker_row = 2

        Dim Lastrow As Long
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
        Dim i As Long

        'Label Ticker Column Headings
        If Need_Summary_Table_Header Then
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
        Else
            Need_Summary_Table_Header = True
        End If
        'open price for the first ticker of current worksheet
        open_price = CurrentWs.Cells(2, 3).Value

        For i = 2 To Lastrow
                     
            If CurrentWs.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'ticker name
                 
                ticker = CurrentWs.Cells(i, 1).Value

                'total volume
                                 
                TotalStockVolume = TotalStockVolume + CurrentWs.Cells(i, 7).Value

                CurrentWs.Range("I" & summary_ticker_row).Value = ticker
                 
                CurrentWs.Range("L" & summary_ticker_row).Value = TotalStockVolume

                close_price = CurrentWs.Cells(i, 6).Value
                Yearly_Change = close_price - open_price

                If open_price <> 0 Then
                    Percent_Change = (Yearly_Change / open_price) * 100

                Else
                End If
                 
                CurrentWs.Range("K" & summary_ticker_row).Value = (CStr(Percent_Change) & "%")

                CurrentWs.Range("J" & summary_ticker_row).Value = Yearly_Change

                If (Yearly_Change > 0) Then
                CurrentWs.Range("J" & summary_ticker_row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    CurrentWs.Range("J" & summary_ticker_row).Interior.ColorIndex = 3
                End If

                summary_ticker_row = summary_ticker_row + 1
                TotalStockVolume = 0

                Yearly_Change = 0
                Percent_Change = 0
                close_price = 0

                open_price = CurrentWs.Cells(i + 1, 3).Value

            Else
                TotalStockVolume = TotalStockVolume + CurrentWs.Cells(i, 7).Value
            End If
        Next i
    Next CurrentWs

End Sub


