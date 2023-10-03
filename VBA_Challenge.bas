Attribute VB_Name = "Module1"
Sub challenge_module_two()

    For Each ws In Worksheets
        
        ' variables
        Dim wsName As String
        Dim tickerName As String
        Dim tickerRow As Integer
        Dim tickerCount As Long
        Dim lastRowColumnA As Long
        Dim lastRowColumnI As Long
        Dim j As Long
        Dim percentChange As Double
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestTotalVolume As Double
        
        ' Getting worksheet name
        wsName = ws.Name
        
        ' Placing the headers for each column
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ' setting starter row
        j = 2
        ' setting ticker count to first row
        tickerCount = 2

        
        ' line to find last Row for column A
        lastRowColumnA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' looping through the data sets
        For i = 2 To lastRowColumnA
                
            ' checks to see if the ticker names change
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' places value of the cell into ticker column
            ws.Cells(tickerCount, 9).Value = ws.Cells(i, 1).Value
            ' places yearly change value into cell
            ws.Cells(tickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            ' this line is to convert the cell values into currency
            ' Range(ws.Cells(j, 10), ws.Cells(i, 10)).Style = "Currency"
            
                ' conditional formating for yearly change
                If ws.Cells(tickerCount, 10).Value < 0 Then
                    
                ' for cells under 0 will have red
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
                    
                Else
            
                ' for cells over 0 will have green
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
                    
                End If
                    
                ' calculating the percent change
                If ws.Cells(j, 3).Value <> 0 Then
                    
                percentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                ' places the value and formats the cells to precent
                ws.Cells(tickerCount, 11).Value = Format(percentChange, "Percent")

                End If
            ' calculating the total volume for the total stock volume column
            ws.Cells(tickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            ' we count the ticker by one
            tickerCount = tickerCount + 1
            ' we add data to the next row in the cell range
            j = i + 1
            
            End If
            
        Next i
        ' using last row I to find the last cell in the this column
        lastRowColumnI = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' initial values for greatest increase, greatest decrease and greatest total volumn
        greatestIncrease = ws.Cells(2, 11).Value
        greatestDecrease = ws.Cells(2, 11).Value
        greatestTotalvolumn = ws.Cells(2, 12).Value
        'looping through the sorted data
        For i = 2 To lastRowColumnI
            
            ' checks data set if the next value is larger and will populate the corresponding cells
            If ws.Cells(i, 11).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
            Else
            
            greatestIncrease = greatestIncrease
            
            End If
            ' checks data set if the next value is smaller and will populate the corresponding cells
            If ws.Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
            Else
            greatestDecrease = greatestDecrease
            
            End If
            ' will find the highest total volumn
            If ws.Cells(i, 12).Value > greatestTotalvolumn Then
            greatestTotalvolumn = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
            Else
            greatestTotalvolumn = greatestTotalvolumn
            
            End If
            
        ' we place the values of total volumn, greatest increase, and greatest decrease with proper formating
        ws.Cells(4, 17).Value = greatestTotalvolumn
        ws.Cells(2, 17).Value = Format(greatestIncrease, "Percent")
        ws.Cells(3, 17).Value = Format(greatestDecrease, "Percent")
        
        Next i
        ' this line allows us to auto fit all the data appropriately within each cell
        Worksheets(wsName).Columns("A:Z").AutoFit
        
            
    Next ws
    
End Sub
