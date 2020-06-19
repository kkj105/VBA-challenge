Attribute VB_Name = "Module1"

Sub StockAnalysis()

    ' Set initial variable for holding the ticker name
    Dim ticker As String
    
    ' Set initial variable for holding the yearly change of price
    Dim yearly_change As Long
    
    ' Set initial variable for holding the percent change over year
    Dim percent_change As Long
    
    ' Set an initial variable for holding the total stock volume
    Dim totalvolume As Double
    totalvolume = 0
     
     ' Set an initial variable for the year opening price of a stock
    Dim year_open As Double
    year_open = Cells(2, 3).Value
    
    ' Set an initial variable for the year closing price of a stock
    Dim year_close As Double
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
     ' Assign header values
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Determine what the last row of data is
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
     
        For i = 2 To LastRow
            ' Check names and change to new cell when ticker name changes
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set the ticker symbol name
            ticker = Cells(i, 1).Value
            
            ' Add to the stock volume
            totalvolume = totalvolume + Cells(i, 7).Value
            
            ' Difference in opening and closing prices
            YearlyChange = Cells(i, 6).Value - year_open
            
            ' Ensure that if year open is 0, then 0 will be reflected in final tally
            If year_open = 0 Then
            PercentChange = 0
            Else
            
            ' Calculate change between year end and year open prices
            PercentChange = YearlyChange / year_open
            End If
            
            ' Check if Yearly Change amount is positive
            If YearlyChange > 0 Then
            
                ' Color the positive amount green
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            ' Check if Yearly Change amount is negative
            ElseIf YearlyChange < 0 Then
                
                ' Color the negative amount red
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
            ' Print the ticker symbol name in the summary table
            Range("I" & Summary_Table_Row).Value = ticker
            
            ' Print the total stock volume in the summary table
            Range("L" & Summary_Table_Row).Value = totalvolume
            
            ' Print the yearly change in the summary table
            Range("J" & Summary_Table_Row).Value = YearlyChange
            
            ' Print the percent change in the summary table
            Range("K" & Summary_Table_Row).Value = PercentChange
            
            ' Format column K to be percentage
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the stock volume total
            totalvolume = 0
                
            'Set the yearly change name
            year_open = Cells(i + 1, 3).Value
            
            ' If the cell immediately following a row is the same ticker symbol, then...
            Else
            
            ' Add to the stock volume
            volume = volume + Cells(i, 7).Value

            End If
        
    Next i

End Sub
