Sub wallstreet()
  
    'Create a loop for each worksheet

    For Each ws In Worksheets
    
        'Activate the current worksheet
        ws.Activate

            'Set a variable for holding the ticker name, YearOpen, YearClose, Yearly_Change, Percent_Change
            Dim Ticker_Name As String
            Dim YearOpen As Double
            Dim YearClose  As Double
            Dim Yearly_Change As Double
            Dim Percent_Change As Double
                
            'Set a variable for holding the total stock volume per ticker
            Dim StockVolume_Total As Double
            StockVolume_Total = 0
                
            'Keep track of the location for each ticker name in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2

            'Populate summary table with titles
            Range("J1").Value = "Ticker"
            Range("K1").Value = "Yearly Change"
            Range("L1").Value = "Percent Change"
            Range("M1").Value = "Total Stock Volume"

            'Set lastRow of dataset
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To LastRow
            
                'Check to see if ticker name is same as next
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                    'Set the ticker name
                    Ticker_Name = Cells(i, 1).Value

                    'Set YearClose value, which is the value from the end date
                    YearClose = Cells(i, 6).Value

                    Yearly_Change = (YearClose - YearOpen)
                    
                    'Add StockVolume to total
                    StockVolume_Total = StockVolume_Total + Cells(i, 7).Value

                    'Calculate Percent Change
                    Percent_Change = Yearly_Change / YearOpen
                    
                    
                    'Print Ticker Name to Summary Table
                    Range("J" & Summary_Table_Row).Value = Ticker_Name

                    'Print Yearly Change  to Summary Table
                    Range("K" & Summary_Table_Row).Value = Yearly_Change

                    'Print Percent Change to Summary Table, as percentage
                    Range("L" & Summary_Table_Row).Value = Percent_Change
                    Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                    'Print StockVolume total to Summary Table
                    Range("M" & Summary_Table_Row).Value = StockVolume_Total
                    
                    'Add to one to summary row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    'Reset Stock Voume to 0
                    StockVolume_Total = 0
                
                'If next ticker name is same as current ticker name
                Else
                    'Add to the stock volume total
                    StockVolume_Total = StockVolume_Total + Cells(i, 7).Value

                    'Need to set YearOpen from earliest date of the year
                    If Cells(i, 2).Value = "20160101" Or Cells(i, 2) = "20150101" Or Cells(i, 2) = "20140101" Then
                        YearOpen = Cells(i, 3).Value
                    End If
                    
                    'Stock ticker of PLNT for 2015 year does not start until August, so must set that value properly
                    If Cells(i, 1).Value = "PLNT" And Cells(i, 3).Value = "0" Then
                        YearOpen = "14.5"
                    End If
                End If
            Next i


            'This solution is for the moderate challenge
            'Set variable to lastrow of the summary table
            Dim sumlastrow As Double

            sumlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

            'Create loop to fill yearly change for each ticker with green (positive) or red color (negative)
            For i = 2 To sumlastrow
                    If (Cells(i, 11).Value > 0) Then
                        Cells(i, 11).Interior.Color = rgbDarkSeaGreen
                        
                    ElseIf (Cells(i, 11).Value < 0) Then
                        Cells(i, 11).Interior.Color = rgbRed
                    End If
            Next i
                
            'This solution is for the the hard challenge
            'Set a variable for max for stockvolume, percent change and min for percent change

            Dim StockVolume_max As Double
            Dim Percent_max As Double
            Dim Percent_min As Double
            

            'Find max of stockvolume, and percent change, and min of percent change and set to variables
                            
            StockVolume_max = WorksheetFunction.Max(Columns("M"))
            Percent_max = WorksheetFunction.Max(Columns("L"))
            Percent_min = WorksheetFunction.Min(Columns("L"))

            'Loop until the script finds the max or min values

            For i = 2 To sumlastrow

                'Populate cells with titles in new summary table
            
                Range("P2").Value = "Greatest % Increase"
                Range("P3").Value = "Greatest % Decrease"
                Range("P4").Value = "Greatest Volume"
                
                Range("Q1").Value = "Ticker"
                Range("R1").Value = "Value"
                
                If Cells(i, 12).Value = Percent_max Then
                    Range("Q2").Value = Cells(i, 10).Value
                    Range("R2").Value = Cells(i, 12).Value
                    Range("R2").NumberFormat = "0.00%"
                End If
                    
                If Cells(i, 12).Value = Percent_min Then
                    Range("Q3").Value = Cells(i, 10).Value
                    Range("R3").Value = Cells(i, 12).Value
                    Range("R3").NumberFormat = "0.00%"
                End If
            
                If Cells(i, 13).Value = StockVolume_max Then
                
                    Range("Q4").Value = Cells(i, 10).Value
                    Range("R4").Value = Cells(i, 13).Value
                    
                End If
            
            
            Next i
    Next ws
    
End Sub
