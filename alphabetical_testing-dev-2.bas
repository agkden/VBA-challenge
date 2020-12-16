Attribute VB_Name = "Module1"
Sub StockMarket()


'-----------------------------------------------
' Declaire and initialize variables
'-----------------------------------------------
    
    ' declaire and set variables
    Dim row_num As Double   ' Integer - not enough
    Dim lastRow As Double   ' Integer - not enough causing overflow error
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox(lastRow)
    
               
    ' set an initial variable for holding the Stock symbol
    Dim ticker As String


    ' declaire and set an initial variable for holding the total volume per ticker
    Dim tickerTotal As Double
    tickerTotal = 0
    
    ' declaire and set variable to keep track of the location for each ticker in the Summary Table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    
     Dim openPrice As Double
     Dim closePrice As Double
     Dim yearlyChange As Double
     Dim percentChange As Double
    
    
    
    
'------------------------------------------
' Create Summary Table Headers
'------------------------------------------
    
    ' add Ticker / Yearly Change / Percent Change / Total Stock Volume as Column Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
   
'--------------------------------------------
    
    
    ' Loop through all <tickers> rows and output
        '  -- Output Ticker Symbol
        '  -- Output Total stock volume
        '  -- Output Yearly Change
        '  -- Output Percent Change
        
    
    'set the first open price value
        openPrice = Cells(2, 3)
    
    For row_num = 2 To lastRow
    
    ' check if we are still within the same ticker name
        ' if it's not
        
            ' tried to put openPrice here   => does'n work as re-write value with each row_num
            '      openPrice = Cells(row_num, 3)
            
            
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


            ' counting percent change
            percentChange = yearlyChange / openPrice
                
            
            ' add to the Summary Table
            Range("K" & summary_table_row).Value = percentChange
                   
        
            ' move to the next row in the Summary Table
            summary_table_row = summary_table_row + 1
            
            
             'set the next open price
            openPrice = Cells(row_num + 1, 3).Value
                
            
            ' Reset Total Volume per ticker
            tickerTotal = 0
            
            
        ' If the cell immediately following a row has the same name
        Else
            
            ' counting a new total per ticker for the one already in Summary Table
            tickerTotal = tickerTotal + Cells(row_num, 7).Value
            
        End If
        
    Next row_num
    
    
End Sub

