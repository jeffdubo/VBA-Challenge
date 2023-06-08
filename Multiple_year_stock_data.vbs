Sub SummarizeStocks():

    ' Cycle through each worksheet and call subroutines to create two tables for each year
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        CreateTickerTable ws
        CreateGreatestTable ws
    
    Next ws
            
End Sub

Sub CreateTickerTable(ws As Worksheet):

    ' Create and format a table summarizing the change in stock value for the year,
    ' the % of change of the stock, and the total stock volume for the year for each ticker/company
    
    Dim i, LastRow As Long
    Dim YearOpen, YearClose, TotalVolume As Double
    
    ' ===========================================
    ' Create the table
    ' ===========================================
        
    ' Create table header
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    ' Initialize YearOpen and TotalVolume for 1st ticker/company
    YearOpen = ws.Range("C2").Value
    TotalVolume = 0
       
    ' Get last row with stock entry for the loop
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Initatize variable to increase row for next ticker/company
    TableRow = 2
    
    ' Loop through every stock entry
    For i = 2 To LastRow

        ' Add current stock volume to total for current ticker/company
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        ' If the next ticker doesn't match, store data for current ticker in the table
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            YearClose = ws.Cells(i, 6).Value
            ws.Cells(TableRow, 9).Value = ws.Cells(i, 1).Value                  ' Store ticker in table
            ws.Cells(TableRow, 10).Value = YearClose - YearOpen                 ' Store change in stock value for the year
            ws.Cells(TableRow, 11).Value = ws.Cells(TableRow, 10) / YearOpen    ' Calculate and store % of change from the opening value
            ws.Cells(TableRow, 12).Value = TotalVolume                          ' Store the total stock volume for the year
            
            YearOpen = ws.Cells(i + 1, 3).Value                                 ' Set the opening value for the year for the next ticker
            TotalVolume = 0                                                     ' Reset total volume for next ticker    
            TableRow = TableRow + 1                                             ' Increase row in table for next ticker
             
        End If
    
    Next i

    ' ===========================================
    ' Part 2: Format the table
    ' ===========================================
    
    TableRow = TableRow - 1                                 ' Reduce by 1 to get the actual last row of the table
    ws.Range("J2:J" & TableRow).NumberFormat = "#.00"       ' Format yearly change data to display 2 decimal places
    ws.Range("K2:K" & TableRow).NumberFormat = "#.00%"      ' Format percent change data as a % with 2 decimal places
    ws.Range("L2:L" & TableRow).NumberFormat = "0"          ' Format total stock volume to prevent scientic notation
    ws.Range("I1:L1").Columns.AutoFit                       ' Autofit width for all columns in the table
    
    ' Remove any existing conditional formating rules
    ws.Range("J2:J" & TableRow).FormatConditions.Delete
    
    ' Create conditional formating rule to set cell color to red for negative yearly changes
    ws.Range("J2:J" & TableRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    ws.Range("J2:J" & TableRow).FormatConditions(1).Interior.Color = vbRed
    
    ' Create conditional formating rule to set cell color to green for positive yearly changes
    ws.Range("J2:J" & TableRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    ws.Range("J2:J" & TableRow).FormatConditions(2).Interior.Color = vbGreen

End Sub

Sub CreateGreatestTable(ws As Worksheet):

    ' Create and format a table with the following info
    '   1. ticker with greatest % increase for the year
    '   2. ticker with greatest % decrease for the year
    '   3. ticker with greatest total stock volume for the year

    Dim i, LastRow As Integer
    Dim TickerMaxInc, TickerMaxDec, TickerMaxVol As String
    Dim MaxInc, MaxDec, MaxVol As Double
    
    ' Set up table with column and row headers
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' Get the last row in the ticker table
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Initialize variables to store the greatest values
    MaxInc = 0
    MaxDec = 0
    MaxVol = 0
   
    ' Loop through every stock entry in the table
    For i = 2 To LastRow
        
        ' If the % change is greater than the current max increase, set the new max and ticker 
        If ws.Cells(i, 11).Value > MaxInc Then
            MaxInc = ws.Cells(i, 11).Value
            IickerMaxInc = ws.Range("I" & i).Value
        
        ' If the % change is less than the current max decrease, set the new max and ticker 
        ElseIf ws.Cells(i, 11).Value < MaxDec Then      '
            IickerMaxDec = ws.Cells(i, 9).Value
            MaxDec = ws.Cells(i, 11).Value
        End If
        
        ' If the total stock volume is greater the current max volumen, set the new max and ticker
        If ws.Cells(i, 12).Value > MaxVol Then
            IickerMaxVol = ws.Cells(i, 9).Value
            MaxVol = ws.Cells(i, 12).Value
        End If
    
    Next i
        
    ' Store the data in the table
    ws.Range("P2").Value = IickerMaxInc
    ws.Range("Q2").Value = MaxInc
    ws.Range("P3").Value = IickerMaxDec
    ws.Range("Q3").Value = MaxDec
    ws.Range("P4").Value = IickerMaxVol
    ws.Range("Q4").Value = MaxVol
    
    ' Format table
    ws.Range("Q2:Q3").NumberFormat = "#.00%"        ' Format greatest precent increase as a % with 2 decimal places
    ws.Range("Q4").NumberFormat = "0"               ' Format total stock volume to prevent scientic notation
    ws.Range("O1:Q4").Columns.AutoFit               ' Autofit width for all columns in the table

End Sub
