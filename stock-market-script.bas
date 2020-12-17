Attribute VB_Name = "Module1"
Sub StockMarket()



    '----------------------------------
    ' Declaire and initialize variables
    '----------------------------------
    
    Dim row_num As Double
    Dim lastRow As Double
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
                
               
    ' set an initial variable for holding the Stock symbol
    Dim ticker As String


    ' declaire and set an initial variable for holding the total volume per ticker
    Dim tickerTotal As Double
    tickerTotal = 0
    
    
    ' declaire and set variable to keep track of the location for each ticker in the Summary Table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    
    ' declaire and set variables for holding open price, close price, yearly change, and percent change
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
     
     
    
    '------------------------------
    ' Create Summary Table Headers
    '------------------------------
    
    ' add Column Headers: Ticker / Yearly Change / Percent Change / Total Stock Volume
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    

    
    '------------------------------------------------
        ' Loop through all <tickers> rows and output
        '  -- Output Ticker Symbol
        '  -- Output Total stock volume
        '  -- Output Yearly Change
        '  -- Output Percent Change
    '-------------------------------------------------
    
    
    'set the first open price value
        openPrice = Cells(2, 3)
    
    
    For row_num = 2 To lastRow
    
    ' check if we are still within the same ticker name
        ' if it's not
        
         If Cells(row_num + 1, 1).Value <> Cells(row_num, 1).Value Then
         
            ' set the ticker
            ticker = Cells(row_num, 1).Value
            
            ' add ticker name to the Summary Table
            Range("I" & summary_table_row).Value = ticker
            
            ' counting a new total per ticker
            tickerTotal = tickerTotal + Cells(row_num, 7).Value
            
            ' add the new total volume to the Summary Table
            Range("L" & summary_table_row).Value = tickerTotal
                
            ' set the close price per ticker
            closePrice = Cells(row_num, 6).Value
                
            ' counting yearly price change
            yearlyChange = closePrice - openPrice
           
            ' add to the Summary Table
            Range("J" & summary_table_row).Value = yearlyChange
 
                ' apply conditional formatting to "Yearly Change" column to highlight
                ' -- positive change in Green (4)
                ' -- negative change in Red (3)
                
                If yearlyChange >= 0 Then
                
                    Range("J" & summary_table_row).Interior.ColorIndex = 4
                
                Else
                
                    Range("J" & summary_table_row).Interior.ColorIndex = 3
                
                End If
                
             
            ' counting percent change
            
                ' set format of Percent Change column in Summary Table to %
                Range("K" & summary_table_row).NumberFormat = "0.00%"
                
                ' check condition for open price
                If openPrice = 0 Then

                    Range("K" & summary_table_row).Value = "0"
                
                Else
                    percentChange = yearlyChange / openPrice
                    
                    Range("K" & summary_table_row).Value = percentChange
                
                End If


            ' move to the next row in the Summary Table
            summary_table_row = summary_table_row + 1
            
            
            ' set the next open price
            openPrice = Cells(row_num + 1, 3).Value
                
                          
            ' reset Total Volume per ticker
            tickerTotal = 0
                
                
            ' If the cell immediately following the current row has the same name
        Else
            
            ' counting a new total per ticker for the symbol already in the Summary Table
            tickerTotal = tickerTotal + Cells(row_num, 7).Value
            
        End If
        
    Next row_num
    

End Sub
