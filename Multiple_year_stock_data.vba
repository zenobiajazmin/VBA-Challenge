Sub stock_analysis():

    ' Set dimensions
    Dim total As Double
    Dim a As Long
    Dim change As Double
    Dim b As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    ' Set initial values
    b = 0
    total = 0
    change = 0
    start = 2

    ' get the row number of the last row with data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For a = 2 To rowCount

        ' If ticker changes then print results
        If Cells(a + 1, 1).Value <> Cells(a, 1).Value Then

            ' Stores results in variables
            total = total + Cells(a, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + b).Value = Cells(a, 1).Value
                Range("J" & 2 + b).Value = 0
                Range("K" & 2 + b).Value = "%" & 0
                Range("L" & 2 + b).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To a
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                change = (Cells(a, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

                ' start of the next stock ticker
                start = a + 1

                ' print the results
                Range("I" & 2 + b).Value = Cells(a, 1).Value
                Range("J" & 2 + b).Value = change
                Range("J" & 2 + b).NumberFormat = "0.00"
                Range("K" & 2 + b).Value = percentChange
                Range("K" & 2 + b).NumberFormat = "0.00%"
                Range("L" & 2 + b).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + b).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + b).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + b).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            b = b + 1
            days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(a, 7).Value

        End If

    Next a

End Sub



