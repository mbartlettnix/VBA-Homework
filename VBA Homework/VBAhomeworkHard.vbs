
'//////////////////////////////////////////////////////////////////////////////////// 
'/Hard
'//Your solution will include everything from the moderate challenge.
'//Your solution will also be able to locate the stock with the "Greatest % increase", 
'//"Greatest % Decrease" and "Greatest total volume".
'////////////////////////////////////////////////////////////////////////////////////
Sub Hard():

Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
        Dim i As LongLong          'to use i as reference'
        Dim rowcounter As LongLong  'output row counter'
        Dim volcounter As Double 'output for unique voume'
        Dim opening As Double   'output for unique opening amount'
        Dim closing As Double   'output for unique closing amount'
        Dim lastrow As LongLong     'identify second number in for i for loop'
        Dim yrchange As Double 'closing-opening'
        Dim perchange As Double
        Dim largest As Double 'largest %'
        Dim smallest As Double 'smallest %'
        Dim yrchangehigh As Double

        rowcounter = 1
        volcounter = 0
        opening = Cells(2, 3).Value
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        largest = 0
        smallest = 0
        yrchangehigh = 0

        'Headers//////////////////////''
        Range("L1").Value = "Ticker"
        Range("M1").Value = "Yearly Change"
        Range("N1").Value = "Percent Change"
        Range("O1").Value = "Total Volume"
        Range("Q2").Value = "Greatest Percent Increase"
        Range("Q3").Value = "Greatest Percent Decrease"
        Range("Q4").Value = "Greatest Volume Change"
        Range("R1").Value = "Ticker"
        Range("S1").Value = "Value"
        Range("Q1:Q4").Borders(xlEdgeLeft).LineStyle = xlContinuous
        Range("S1:S4").Borders(xlEdgeRight).LineStyle = xlContinuous
        Range("Q1:S1").Borders(xlEdgeTop).LineStyle = xlContinuous
        Range("Q4:S4").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Columns(13).AutoFit
        Columns(14).AutoFit
        Columns(15).AutoFit
        Columns(17).AutoFit
        Columns(17).AutoFit
        '//////////////////////////////'


        For i = 2 To lastrow
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
            closing = Cells(i, 6).Value
            yrchange = closing - opening
            If opening = 0 Then 'To work past stocks starting with 0'
                perchange = 0            
                Else
                perchange = (1 - (closing / opening)) * 100
                End If
            Cells(1 + rowcounter, 12).Value = Cells(i, 1).Value
            Cells(1 + rowcounter, 15).Value = Cells(i, 7).Value + volcounter
            Cells(1 + rowcounter, 13).Value = yrchange
                If yrchange < 0 Then
                Cells(1 + rowcounter, 13).Interior.ColorIndex = 3
                Else: Cells(1 + rowcounter, 13).Interior.ColorIndex = 43
                End If
            
                If yrchange > yrchangehigh Then
                yrchangehigh = yrchange
                Range("R4").Value = Cells(i, 1).Value
                Range("S4").Value = yrchangehigh
                End If

            Cells(1 + rowcounter, 14).Value = perchange & "%"
                If perchange > largest Then
                largest = perchange
                Range("R2").Value = Cells(i, 1).Value
                Range("S2").Value = largest & "%"
                ElseIf perchange < smallest Then
                smallest = perchange
                Range("R3").Value = Cells(i, 1).Value
                Range("S3").Value = smallest & "%"
                End If

            opening = Cells(i + 1, 3).Value

            rowcounter = rowcounter + 1
            volcounter = 0

            Else
            Cells(1 + rowcounter, 12).Value = Cells(i, 1).Value
            volcounter = Cells(i, 7).Value + volcounter
            End If
        Next i
    Next ws
End Sub
