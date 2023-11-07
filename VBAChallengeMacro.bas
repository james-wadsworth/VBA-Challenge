Attribute VB_Name = "Module1"
Sub stock_ticker()

    'Loop this through all worksheets
    For Each ws In Worksheets
    
        'Set up the column and row headers properly for the summary areas
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set these up for the second summary area
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
            'Create variables to track the values to be summarized -
            'Ticker symbol, price change,percent change, and volume
            Dim Ticker As String
            Dim PercentChange As Double
            Dim PriceChange As Double
            Dim Volume As Double
            
            
            'Set variable values to 0 for variables tracking numbers
            PercentChange = 0
            Volume = 0
            PriceChange = 0
            
                'For PriceChange variable, to track price change, create a variable to hold the value of the cell with the first open of the year in it
                Dim OpenPrice As Double
                
                'Set the initial open price
                OpenPrice = ws.Cells(2, 3).Value
                
                'Create a variable to hold the last closing price and set it to 0 also
                Dim ClosingPrice As Double
                ClosingPrice = 0
            
            'Create a variable to store the last row number and set it equal to the last row
            Dim lastRow As Integer
            lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
            
            'Create a row to hold the number of the row in which to print summary info next and start it at first row of summary area
            Dim SummaryTableRow As Integer
            SummaryTableRow = 2
            
            'Set up the loop to go through the values in column 1 and find where ticker changes then print summary info
            For i = 2 To lastRow
            
                
                'At each cell, check if the next ticker value is different
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    'Set the Ticker symbol
                    Ticker = ws.Cells(i, 1).Value
                
                    'Print the name of the ticker in the first row of the summary table
                    ws.Cells(SummaryTableRow, 9).Value = Ticker
                    
                    'Set the volume
                    Volume = Volume + ws.Cells(i, 7).Value
                    
                    'Print the volume in the summary table
                    ws.Cells(SummaryTableRow, 12).Value = Volume
                    
                    'Set the Closing Price
                    ClosingPrice = ws.Cells(i, 6).Value
                    
                    'Calculate the Price Change
                    PriceChange = ClosingPrice - OpenPrice
                    
                    'Print the Price Change
                    ws.Cells(SummaryTableRow, 10).Value = PriceChange
                    
                    'Calculate the percent change
                    PercentChange = PriceChange / OpenPrice
                    
                    'Print the percent change
                    ws.Cells(SummaryTableRow, 11).Value = PercentChange
                    
                    'Increment the summary table row
                    SummaryTableRow = SummaryTableRow + 1
                    
                    'Reset the volume
                    Volume = 0
                    
                    'Reset the open price
                    OpenPrice = ws.Cells(i + 1, 3).Value
                    
                Else
                
                    'Add to the volume total
                    Volume = Volume + ws.Cells(i, 7).Value
                    
                End If
                
            Next i
            
            'Find the greatest increase, decrease, and total volume for the year
            
            
            
                'Find the last row in the summary table
                Dim SummaryLastRow As Integer
                SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                
                'Find the max value and hold it in a variable
                Dim GreatestIncrease As Double
                GreatestIncrease = WorksheetFunction.Max(ws.Range("K2:K" & SummaryLastRow))
                        
                'Loop through the values in column K or the Percent Change column to find the Ticker
                For i = 2 To SummaryLastRow
                
                    'Check at each cell if this is the max
                   If ws.Cells(i, 11).Value = GreatestIncrease Then
                   
                        'Print the Ticker and the max in the correct cells
                        ws.Cells(2, 16).Value = Cells(i, 1).Value
                        ws.Cells(2, 17).Value = GreatestIncrease
                    End If
                    
                        
                        'Format the cells as you go based on positive (green) or negative(red)
                        Set condition1 = ws.Cells(i, 11).FormatConditions.Add(xlCellValue, xlGreater, "=0")
                        Set condition2 = ws.Cells(i, 11).FormatConditions.Add(xlCellValue, xlLess, "=0")
                        
                        With condition1
                            .Interior.ColorIndex = 4
                        End With
                        With condition2
                            .Interior.ColorIndex = 3
                        End With
                        
                        'Format the cells in the same conditional manner for the "Yearly Change" column as you go based on positive (green) or negative(red)
                        'Format the cells as you go based on positive (green) or negative(red)
                        Set condition1 = ws.Cells(i, 10).FormatConditions.Add(xlCellValue, xlGreater, "=0")
                        Set condition2 = ws.Cells(i, 10).FormatConditions.Add(xlCellValue, xlLess, "=0")
                        
                        With condition1
                            .Interior.ColorIndex = 4
                        End With
                        With condition2
                            .Interior.ColorIndex = 3
                        End With
                        
                        
                Next i
                
                
                'Find the min value and hold it in a variable
                Dim GreatestDecrease As Double
                GreatestDecrease = WorksheetFunction.Min(ws.Range("K2:K" & SummaryLastRow))
                        
                'Loop through the values in column K or the Percent Change column to find the Ticker
                For i = 2 To SummaryLastRow
                
                    'Check at each cell if this is the max
                   If ws.Cells(i, 11).Value = GreatestDecrease Then
                   
                        'Print the Ticker and the max in the correct cells
                        ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
                        ws.Cells(3, 17).Value = GreatestDecrease
                    End If
                Next i
                
                
                'Find the max volume and hold it in a variable
                Dim MaxVolume As Double
                MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & SummaryLastRow))
                
                'Loop through the values in column K or the Percent Change column to find the Ticker
                For i = 2 To SummaryLastRow
                
                    'Check at each cell if this is the max
                   If ws.Cells(i, 12).Value = MaxVolume Then
                   
                        'Print the Ticker and the max in the correct cells
                        ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
                        ws.Cells(4, 17).Value = MaxVolume
                    End If
                Next i
                
                
              
     Next ws
            
End Sub
