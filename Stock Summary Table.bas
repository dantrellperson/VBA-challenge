Attribute VB_Name = "Module1"
'This code will create a summary table for the provided stock market data

Sub stocks():
    'variable to hold ticker
    ticker = " "
    
    'variable to hold total stock volume for ticker
    totalStockVolume = 0
    
    'variable to hold close price ticker
    Dim closePrice As Double
    
    'variable to hold open price for ticker
    Dim openPrice As Double
    openPrice = Cells(2, 3).Value
    
    'variable to hold yearly change
    yearlyChange = 0
    
    'variable to hold percent change
    percentChange = 0
    
    'variable to hold the summary table start row
    summaryTableRow = 2
        
    'use last row function
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop from row 2 in column A out to the last row
    For Row = 2 To lastrow
        
        'check to see if the ticker changes
        
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
            
                'if the brand changes, do...
                
                'first set the ticker
                ticker = Cells(Row, 1).Value
                
                'set the closeprice
                closePrice = Cells(Row, 6).Value
                 
                'add the last volume from that row
                totalStockVolume = totalStockVolume + Cells(Row, 7).Value
                
                'calculate yearlyChange
                yearlyChange = closePrice - openPrice
                
                'calculate percentChange
                percentChange = (yearlyChange / openPrice)
                percentChange = Format(percentChange, "0.00%")
                                              
                'add the ticker to the i column starting in the summary table row
                Cells(summaryTableRow, 9).Value = ticker
                
                'add the yearlyChange to the J column in the summary table row
                Cells(summaryTableRow, 10).Value = yearlyChange

                'add the percentChange to the K column in the summary table row
                Cells(summaryTableRow, 11).Value = percentChange
                
                'add the totalStockVolume to the L column in the summary table row
                Cells(summaryTableRow, 12).Value = totalStockVolume
                
                'go to the next summary table row (add 1 on to the value of the summary table row)
                summaryTableRow = summaryTableRow + 1
                
                'reset the total sock volume to 0
                totalStockVolume = 0

                'reset the yearly change to 0
                yearlyChange = 0
                
                'reset open price
                openPrice = Cells(Row + 1, 3).Value
                
                'reset close price
                closePrice = 0
                
            Else
         
                'if the ticker stays the same, do
                'add on to the Total Stock Volume from the G column
                totalStockVolume = totalStockVolume + Cells(Row, 7).Value
                
            End If
                         
    Next
    
    'Now its time to color!
    
    For Row = 2 To lastrow
        
        For Column = 10 To 10
        
            If Cells(Row, Column).Value > 0 Then
            
                Cells(Row, Column).Interior.ColorIndex = 4 'Green for positive
                
            ElseIf Cells(Row, Column).Value < 0 Then
            
                Cells(Row, Column).Interior.ColorIndex = 3 'Red for negative
                
             End If
             
        Next Column
        
    Next Row
        
        
        
        
    
      
End Sub



