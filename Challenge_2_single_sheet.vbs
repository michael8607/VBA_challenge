Sub stock_analysis():
' Set dimensions
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Integer
Dim start As Long
Dim row_count As Long
Dim percentChange As Double
Dim days As Integer
Dim dailyChange As Double
Dim averageChange As Double


' Set title
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

' Set initial values
j = 0
total = 0
change = 0
start = 2

' get the last row with data
row_count = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To row_count

' If ticker changes from one row to the next then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ' Store results in variable total
        total = total + Cells(i, 7).Value
    ' Handle in case of zero total volume
        If total = 0 Then
        ' print results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
        Else
        ' Find First starting value greater than 0
            If Cells(start, 3) = 0 Then
                For find_value = start To i
                    If Cells(find_value, 3).Value <> 0 Then
                        start = find_value
                        Exit For
                    End If
                Next find_value
            End If

        ' Calculate Change between values
            change = (Cells(i, 6) - Cells(start, 3))
            percentChange = change / Cells(start, 3)
        ' start of the next stock
            start = i + 1
        ' print results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = change
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = percentChange
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Range("L" & 2 + j).Value = total

        ' format positives & negatives in respective colors
            Select Case change
                Case Is > 0
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
        End If
    ' reset variables for new stock (increase j by 1 to add new row
        total = 0
        change = 0
        j = j + 1
        days = 0
    ' If ticker is the same add results together
    Else
        total = total + Cells(i, 7).Value
    End If
Next i

' take the max and min
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & row_count)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & row_count)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & row_count))

' returns one less to account for header row not being a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & row_count)), Range("K2:K" & row_count), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & row_count)), Range("K2:K" & row_count), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & row_count)), Range("L2:L" & row_count), 0)

' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)


End Sub