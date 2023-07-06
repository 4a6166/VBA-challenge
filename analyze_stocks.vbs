Sub clearTables()
    ' clears cols i-p to help with testing
    For Each ws In Sheets
        ws.Activate
        Range("i1:p500").Clear
        Range("i1:p500").FormatConditions.Delete
        
    Next ws
End Sub

Sub runAnalysis()

    For Each ws In Sheets
        ' set each ws as active, don't pass ws to subs
        ws.Activate
        
        setupTables
        sumarizeStocks
        addConditionals
        
    Next ws
End Sub

Sub sumarizeStocks()

   Dim tickers() As String
   Dim dateLow() As Variant
   Dim dateHigh() As Variant
   Dim volumes() As LongLong
   ' Dates stored as strings in data. Apparently arithmetic still works.
   Dim firstOpens() As String
   Dim lastCloses() As String

   ReDim tickers(0)
   ReDim firstOpens(0)
   ReDim lastCloses(0)
   ReDim volumes(0)
   
   Dim valueGreatestIncrease As Double
   Dim valueGreatestDecrease As Double
   Dim valueGreatestVolume As LongLong
   
    ' lastrow algorithm provided by instructor via slack
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 2 To lastrow
    
        Dim record(6) As Variant

        record(0) = Cells(Row, 1).Value 'ticker
        record(1) = Cells(Row, 2).Value 'date
        record(2) = Cells(Row, 3).Value 'open
        record(3) = Cells(Row, 4).Value 'high
        record(4) = Cells(Row, 5).Value 'low
        record(5) = Cells(Row, 6).Value 'close
        record(6) = Cells(Row, 7).Value 'volume
    
        SumTableRow = 0
        
        ' search sum table for tracked tickers
        For i = 0 To UBound(tickers)
            If tickers(i) = record(0) Then
                SumTableRow = i
                Exit For
            End If
        Next i
        
        ' add untracked ticker to tickers array, create space in open, close, vol arrays
        If SumTableRow = 0 Then
            newUpper = UBound(tickers) + 1
            ReDim Preserve tickers(newUpper)
            ReDim Preserve dateLow(newUpper)
            ReDim Preserve dateHigh(newUpper)
            ReDim Preserve firstOpens(newUpper)
            ReDim Preserve lastCloses(newUpper)
            ReDim Preserve volumes(newUpper)
            
            tickers(newUpper) = record(0)
            dateLow(newUpper) = record(1)
            dateHigh(newUpper) = record(1)
            firstOpens(newUpper) = record(2)
            lastCloses(newUpper) = record(5)
            volumes(newUpper) = record(6)
            
            SumTableRow = UBound(tickers)
            
        Else
            If record(1) < dateLow(SumTableRow) Then
                dateLow(SumTableRow) = record(1)
                firstOpens(SumTableRow) = record(2)
            End If
            
            If record(1) > dateHigh(SumTableRow) Then
                dateHigh(SumTableRow) = record(1)
                lastCloses(SumTableRow) = record(5)
            End If
            
            volumes(SumTableRow) = volumes(SumTableRow) + record(6)
        End If
   
    Next Row
    
    ' workaround for type error in valueGreatstIncrease leading to incorrect value
    ' can't instantiate as a decimal?
    valueGreatestIncrease = (lastCloses(i) - firstOpens(i)) / firstOpens(i)

    For i = LBound(tickers) + 1 To UBound(tickers)
        ' Print summary arrays to table
        Offset = i + 1
        annualChange = lastCloses(i) - firstOpens(i)
        percentChange = (annualChange) / firstOpens(i)
        
        Cells(Offset, 9).Value = tickers(i)
        Cells(Offset, 10).Value = annualChange
        Cells(Offset, 11).Value = percentChange
        Cells(Offset, 12).Value = volumes(i)
        
        ' Print greatest values
        If valueGreatestIncrease < percentChange Then
            valueGreatestIncrease = percentChange
            Cells(2, 15).Value = tickers(i)
            Cells(2, 16).Value = percentChange
        End If
        
        If valueGreatestDecrease > percentChange Then
            valueGreatestDecrease = percentChange
            Cells(3, 15).Value = tickers(i)
            Cells(3, 16).Value = percentChange
        End If
        
        If valueGreatestVolume < volumes(i) Then
            valueGreatestVolume = volumes(i)
            Cells(4, 15).Value = tickers(i)
            Cells(4, 16).Value = volumes(i)
        End If
        
    Next i

    
End Sub

Sub setupTables()

    ' Tickers totals
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    
    'Summary table
    Range("o1").Value = "Ticker"
    Range("p1").Value = "Value"
    
    Range("n2").Value = "Greatest % Increase"
    Range("n3").Value = "Greatest % Decrease"
    Range("n4").Value = "Greatest Total Volume"
    
    'Format cells
    Range("k:k, p2:p3").Style = "Percent"
    Range("k:k, p2:p3").NumberFormat = "0.00%"
    
    Range("l:l, p4").NumberFormat = "0"

End Sub

Sub addConditionals()
    'remove any current formatting
    Columns("j:k").FormatConditions.Delete
    
    ' add conditional formatting for less than 0
    Set YCLess = Range("j:k").FormatConditions.Add(xlCellValue, xlLess, "=0")
    With YCLess
        .Interior.Color = vbRed
    End With
    
    ' add conditional formatting for less than 0
    Set YCGreater = Range("j:k").FormatConditions.Add(xlCellValue, xlGreater, "=0")
    With YCGreater
        .Interior.Color = vbGreen
    End With
    
    Range("j1:k1").FormatConditions.Delete

End Sub
