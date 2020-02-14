Sub challenge1():

    Dim total As Double

    totalRows = Cells(Rows.Count, "A").End(xlUp).Row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    For row = 2 To totalRows

        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then

            total = total + Cells(row, 7).Value

            Range("I" & 2 + CurrentCell).Value = Cells(row, 1).Value

            Range("J" & 2 + CurrentCell).Value = total

            total = 0

            CurrentCell =  CurrentCell + 1

        Else
            total = total + Cells(row, 7).Value

        End If

    Next row

End Sub