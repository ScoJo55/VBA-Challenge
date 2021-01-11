Attribute VB_Name = "Module1"
Sub StockData()
'loop to all worksheets in workbook

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
    
'define variable
Dim yropen As Double
Dim yrclose As Double
Dim f As Integer
Dim a As Long
Dim i As Long
Dim j As Long
Dim ws_count As Integer
Dim total As LongLong

'add headers and titles
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "% change"
Cells(1, 12) = "Total Stock Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"

'format columns
Range("K:K").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"
Range("O:O").ColumnWidth = 20
Range("Q:Q").ColumnWidth = 15
Range("L:L").ColumnWidth = 20
Range("J:J").ColumnWidth = 15

'define the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
total = 0

'loop for ticker and adding the total volume
 f = 2
    For c = 2 To lastrow
   
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
            Cells(f, 9) = Cells(c, 1).Value
            total = total + Cells(c, 7).Value
            Cells(f, 12) = total
            f = f + 1
            total = 0
        Else
            total = total + Cells(c, 7).Value
        End If
        
    Next c

yropen = Cells(2, 3).Value

'loop for change
    a = 2
        For i = 2 To lastrow
    
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                yrclose = Cells(i, 6).Value
                Cells(a, 10).Value = yrclose - yropen
            If yropen = 0 Then
                Cells(a, 11).Value = 0
                yropen = Cells(i + 1, 3).Value
                a = a + 1
            Else
                Cells(a, 11).Value = Cells(a, 10).Value / yropen
                yropen = Cells(i + 1, 3).Value
                a = a + 1
                
            End If
            End If
        Next i

'define last row for display cells
dislastrow = Cells(Rows.Count, 9).End(xlUp).Row
    
'loop to add color
    For j = 2 To dislastrow
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        
        End If
    Next j
    
'define variables
Dim rng As Range
Dim rng2 As Range
Dim min As Double
Dim max As Double
Dim volmax As LongLong
    
'set the ranges to calculate min and max change and max total volume
    Set rng = Range("K:K")
    Set rng2 = Range("L:L")
    
'calculates the min and max and reports in column Q
    max = WorksheetFunction.max(rng)
    min = WorksheetFunction.min(rng)
    volmax = WorksheetFunction.max(rng2)
    Range("Q2").Value = max
    Range("Q3").Value = min
    Range("Q4").Value = volmax
   

    For k = 2 To dislastrow
        If Cells(k, 11).Value = max Then
            Cells(2, 16) = Cells(k, 9).Value
           

        End If
    Next k
    
'loop for min value and return the ticker sign
    For g = 2 To dislastrow
        If Cells(g, 11).Value = min Then
            Cells(3, 16) = Cells(g, 9).Value
    
        End If
    Next g
    
'loop maxtotal volume value and return the ticker sign
    For m = 2 To dislastrow
        If Cells(m, 12).Value = volmax Then
            Cells(4, 16) = Cells(m, 9).Value
        
        End If
    Next m


'next worksheet
Next ws
        
    
    
    
End Sub
