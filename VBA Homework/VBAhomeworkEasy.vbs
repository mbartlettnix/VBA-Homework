'////////////////////////////////////////////////////////////////////////////////////
'// Easy
'// Create a script that will loop through each year of stock data and grab the total 
'// amount of volume each stock had over the year.
'// You will also need to display the ticker symbol to coincide with the total volume.
'////////////////////////////////////////////////////////////////////////////////////
	
Sub Easy():
 
Dim ws As Worksheet
    
   For Each ws In Worksheets
        ws.Activate
        Dim i As Long
        Dim rowcounter As Long
        Dim volcounter As Double
        
        Dim lastrow As Long
        
        rowcounter = 0
        volcounter = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Headers
       Range("L1").Value = "Ticker"
       Range("M1").Value = "Total Volume"
        
        For i = 2 To lastrow
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
            Cells(1 + rowcounter, 12).Value = Cells(i, 1).Value
            Cells(1 + rowcounter, 13).Value = Cells(i, 7).Value + volcounter
            rowcounter = rowcounter + 1
            volcounter = 0
            Else
            Cells(1 + rowcounter, 12).Value = Cells(i, 1).Value
            volcounter = Cells(i, 7).Value + volcounter
            End If
        Next i
    Next ws
End Sub







