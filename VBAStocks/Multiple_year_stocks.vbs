Sub multiple_year_stock():
    
    Dim sht As Worksheet
    For Each sht In Worksheets
        sht.Activate
        
        'declare variables
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim YearOpen As Double
        Dim YearClose As Double
   
        'create header for summary_table with columns for: ticker, yearly change, percent change, total stock volume
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change "
        Range("L1").Value = "Total Volume"
    
        'define last row
        Dim Last_Row As Long
        Last_Row = sht.Cells(Rows.Count, 1).End(xlUp).Row
    
        'define summary_table
        Dim summary_table As Integer
        
        'prepere for the loop by assigning "0" value to the total volume and starting the summary_table from 2
        summary_table = 2
        TotalVolume = 0
    
    
        'set up  Open year
        YearOpen = Range("C2").Value
    
        'loop through the sheet and retrive the data in the summary_table
        For i = 2 To Last_Row

            'search through the ticker column for diferent tickers and add the total volume
            If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
        
                Ticker = Cells(i, 1).Value
            
                TotalVolume = TotalVolume + Cells(i, 7).Value
            
                'print the ticker symbols
                Range("I" & summary_table).Value = Ticker
            
                'print total volume
                Range("L" & summary_table).Value = TotalVolume
                
                YearClose = Range("F" & i).Value
            
                'calculate yearly change by substarcting value of YearOpen from YearClose
                YearlyChange = YearClose - YearOpen
            
                'print the value of yearly change
                Range("J" & summary_table).Value = YearlyChange
            
                'color accordingly based on negative and positive change in yearly change column by usuing the color index
                    If (YearlyChange > 0) Then
                        Range("J" & summary_table).Interior.ColorIndex = 4
                    Else
                        Range("J" & summary_table).Interior.ColorIndex = 3
                    End If
                    
                'calculate percent change
                PercentChange = (YearClose - YearOpen) / YearOpen
                
                'print the value of percent change
                Range("K" & summary_table).Value = PercentChange
                
                'format to percentage to add the "%" sign using Number Format
                Range("K" & summary_table).NumberFormat = "0.00%"
            
                summary_table = summary_table + 1
            
                ' reset the value of total volume
                TotalVolume = 0
                YearInitial = Range("C" & i + 1).Value
            
            Else
                TotalVolume = TotalVolume + Cells(i, 7).Value
         
        End If
            
    Next i

Next sht


End Sub
