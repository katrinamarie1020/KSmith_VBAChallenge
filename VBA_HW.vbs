Sub HW()

Dim current_ticker As String
Dim next_ticker As String
Dim totalrows As Long
Dim total As Long
Dim ticker_tablerow As Long

totalrows = Cells(Rows.Count, "A").End(xlUp).Row

ticker_tablerow = 2

For i = 2 To totalrows

current_ticker = Cells(i, 1).Value
next_ticker = Cells(i + 1, 1).Value

total = Clngtotal + Cells(i, 7).Value

If current_ticker <> next_ticker Then
    Cells(ticker_tablerow, 9).Value = current_ticker
    Cells(ticker_tablerow, 11).Value = total

ticker_tablerow = ticker_tablerow + 1

total = 0

End If

Next i
End Sub