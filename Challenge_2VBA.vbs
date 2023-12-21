Sub stock_analysis():
' Set dimensions
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Integer
Dim start As Long
Dim row_count As Long
Dim percent_change As Double
Dim days As Integer
Dim dailyChange As Double
Dim averageChange As Double
Dim ws As Worksheet

' loop thru all worksheets

For Each ws In Worksheets
' Set title
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

' Set initial values
j = 0
total = 0
change = 0
start = 2

' get the last row with data
row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To row_count

' If changes from one row to the next then
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ' Store results in variable total
        total = total + ws.Cells(i, 7).Value
    ' Handle in case of zero total volume
        If total = 0 Then
        ' print results
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = 0
            ws.Range("K" & 2 + j).Value = "%" & 0
            ws.Range("L" & 2 + j).Value = 0
        Else
        ' Find First value greater than 0
            If ws.Cells(start, 3) = 0 Then
                For find_value = start To i
                    If ws.Cells(find_value, 3).Value <> 0 Then
                        start = find_value
                        Exit For
                    End If
                Next find_value
            End If

        ' Calculate change 
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            percent_change = change / ws.Cells(start, 3)
        ' start the next stock
            start = i + 1
        ' print 
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = change
            ws.Range("J" & 2 + j).NumberFormat = "0.00"
            ws.Range("K" & 2 + j).Value = percent_change
            ws.Range("K" & 2 + j).NumberFormat = "0.00%"
            ws.Range("L" & 2 + j).Value = total

        ' format positives & negatives in respective colors
            Select Case change
                Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
        End If
    ' reset variables for new stock (increase j by 1 to add new row
        total = 0
        change = 0
        j = j + 1
        days = 0
    ' If ticker is the same add results together
    Else
        total = total + ws.Cells(i, 7).Value
    End If
Next i

' max and min
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & row_count)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & row_count)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & row_count))

' returns one less to account for header row not being a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & row_count)), ws.Range("K2:K" & row_count), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & row_count)), ws.Range("K2:K" & row_count), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & row_count)), ws.Range("L2:L" & row_count), 0)

' final total ticker #, greatest diff between increase vs decrease, and the average
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)
Next ws

End Sub