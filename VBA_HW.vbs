Sub HW_2()

Dim current_ticker As String
Dim next_ticker As String
Dim totalrows As Double
Dim total As Double
Dim ticker_tablerow As Double
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double

open_price = Range("C2").Value
totalrows = Cells(Rows.Count, "A").End(xlUp).Row

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Change"

ticker_tablerow = 2

For i = 2 To totalrows

    current_ticker = Cells(i, 1).Value
    next_ticker = Cells(i + 1, 1).Value


    total = Clngtotal + Cells(i, 7).Value

        If current_ticker <> next_ticker Then
            Cells(ticker_tablerow, 9).Value = current_ticker
            Cells(ticker_tablerow, 12).Value = total
            close_price = Cells(i, 6).Value
            
            Cells(ticker_tablerow, 10) = (close_price - open_price)
            
            percent_change = (close_price - open_price) / open_price
            Cells(ticker_tablerow, 11) = percent_change
            
            Cells(ticker_tablerow, 11).NumberFormat = "0.0%"
    
            If percent_change > 0 Then
                Cells(ticker_tablerow, 11).Interior.ColorIndex = 4
        
            Else: Cells(ticker_tablerow, 11).Interior.ColorIndex = 3
            End If
                    
            ticker_tablerow = ticker_tablerow + 1
 
            total = 0


        End If

Next i

End Sub