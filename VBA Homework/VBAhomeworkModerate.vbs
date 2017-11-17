'//////////////////////////////////////////////////////////////////////////////////// 
'// Moderate
'// Create a script that will loop through all the stocks and take the following info:
'//// Yearly change from what the stock opened the year at to what the closing price was.
'//// The percent change from what it opened the year at to what it closed.
'//// The total Volume of the stock
'//// Ticker symbol(Name)
'//// You should also have conditional formatting that will highlight positive change 
'///// in green and negative change in red.
'////////////////////////////////////////////////////////////////////////////////////

Sub Moderate():
 
Dim ws As Worksheet
    
   For Each ws In Worksheets
        ws.Activate
        Dim i As Long           'to use i as reference'
        Dim rowcounter As Long  'output row counter'
        Dim volcounter As Double 'output for unique voume'
        Dim opening As Double   'output for unique opening amount'
        Dim closing As Double   'output for unique closing amount'
        Dim lastrow As Long     'identify second number in for i for loop'
        Dim yrchange As Double 'closing-opening'
        
        rowcounter = 1
        volcounter = 0
        opening = Cells(2, 3).Value
       
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
       'Headers
       Range("L1").Value = "Ticker"
       Range("M1").Value = "Yearly Change"
       Range("N1").Value = "Percent Change"
       Range("O1").Value = "Total Volume"
       
        For i = 2 To lastrow
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
            closing = Cells(i, 6).Value
            yrchange = closing - opening
            
            Cells(1 + rowcounter, 12).Value = Cells(i, 1).Value
            Cells(1 + rowcounter, 15).Value = Cells(i, 7).Value + volcounter
            Cells(1 + rowcounter, 13).Value = yrchange
                If yrchange < 0 Then
                Cells(1 + rowcounter, 13).Interior.ColorIndex = 3
                Else
                Cells(1 + rowcounter, 13).Interior.ColorIndex = 43
                End If
            Cells(1 + rowcounter, 14).Value = (1- (closing / opening)) * 100 & "%"

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



