Sub Looptest2()
    
    'Create column titles for where summary data will be placed
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change Open to Close"
    Cells(1, 11).Value = "Yearly Percent Change Open to Close"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Create a variable to hold Tickercounter
    Dim Tickercounter As Integer
    Tickercounter = 0
    
    'Create a variable to hold YearlyChange
    Dim Yearlychange As Double
    Yearlychange = 0
    
    'Create a variable to hold YearlyPercentChange
    Dim YearlyPercentChange As Double
    YearlyPercentChange = 0
    
    'Create a variable to hold Totalstockvolume
    Dim totalstockvolume As Double
    totalstockvolume = 0
    
    'Count the number of rows
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' first row
    firstrow = 2
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'Loop through each row, use last row instead of actual row number count
    For i = 2 To lastrow
    
        'Keep track of the location for each ticker in the summary table
​
        
        Ticker = Cells(i, 1).Value
        tickerdate = Cells(i, 2).Value
        tickeropen = Cells(firstrow, 3).Value
        tickerclose = Cells(i, 6).Value
        stockvalue = Cells(i, 7).Value
             
        If Cells(i + 1, 1).Value <> Ticker Then
            'Calculate change from open to close
            Yearlychange = tickerclose - tickeropen
            'Calculate percent change from open to close
            YearlyPercentChange = (Yearlychange / tickeropen) * 100
            
            totalstockvolume = totalstockvolume + Cells(i, 7).Value
            'Print the ticker in the summary table
            Range("I" & Summary_Table_Row).Value = Ticker
             
            'Print the YearlyChange in the summary table
            Range("J" & Summary_Table_Row).Value = Yearlychange
                
            'Print the YearlyPercentChange in the summary table
            Range("K" & Summary_Table_Row).Value = YearlyPercentChange
            
            'Print the totalstockvolume in the summary table
            Range("L" & Summary_Table_Row).Value = totalstockvolume
​
​
            firstrow = i + 1
            
                        
            'Reset Tickercounter
            Tickercounter = 0
            
            'Reset yearlychange
            Yearlychange = 0
            
            'Reset yearlypercenchange
            YearlyPercentChange = 0
            
            'Reset the totalstockvolume
            totalstockvolume = 0
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
         'If the cell immediately following a row is the same ticker
        Else
            totalstockvolume = totalstockvolume + Cells(i, 7).Value
        End If
    Next i
​
End Sub
  
