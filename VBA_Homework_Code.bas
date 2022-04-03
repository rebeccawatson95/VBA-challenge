Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data():

For Each ws In Worksheets
ws.Activate


'add new column headers:
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Value"

'set data to type:
Dim j As Integer
Dim total As Double
Dim change As Double
Dim start As Long
Dim i As Long
Dim last_line As Long
Dim percent_change As Double
Dim days As Integer


'set initial values:
j = 0
total = 0
change = 0
start = 2

'run code to last line containing data:
last_line = Cells(Rows.Count, "A").End(xlUp).Row

'start looping through data:
For i = 2 To last_line
    'if ticker changes, print cells
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        total = total + Cells(i, 7).Value
        'if zero change, print results
        If total = 0 Then
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
        'now find new start
        Else
            If Cells(start, 3) = 0 Then
                For new_start = start To i
                    If Cells(new_start, 3).Value <> 0 Then
                        start = new_start
                        Exit For
                    End If
                Next new_start
            End If
            'calculate yearly and percent change
            change = (Cells(i, 6) - Cells(start, 3))
            percent_change = change / Cells(start, 3)
            'next ticker
            start = i + 1
            'print results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = change
            Range("K" & 2 + j).Value = percent_change
            Range("L" & 2 + j).Value = total
            Range("j" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).NumberFormat = "0.00%"
            'color cells based on positive vs negative
            Select Case change
                Case Is > 0
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
        End If
        'reset for next ticker
        total = 0
        change = 0
        j = j + 1
        days = 0
    Else
        total = total + Cells(i, 7).Value
    End If
Next i
Next ws
End Sub

