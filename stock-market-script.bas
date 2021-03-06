Attribute VB_Name = "Module2"
Sub StockMarket()

    ' Declare ws as a worksheet object variable
    Dim ws As Worksheet

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
    ' Loop through all of the worksheets in the active workbook
    For Each ws In Worksheets
        'MsgBox ws.Name
        

        '----------------------------------
        ' Declaire and initialize variables
        '----------------------------------
        
        Dim row_num As Double
        Dim lastRow As Double
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (lastRow)
            
                       
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
            
        ' add Column Headers: Ticker | Yearly Change | Percent Change | Total Stock Volume
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
        ' Create Headers for columns and rows of "Bonus part" Summary Table
        ' add Column names: Ticker | Value to columns 'P' and 'Q'
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' add Row names
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' change column and row headers of Summary Table to Bold type
        ws.Range("I1:Q1").Font.Bold = True
        ws.Range("O2:O4").Font.Bold = True
        
        ' add autofit column width formatting to the Headers
        ' ws.Range("I:Q").Columns.AutoFit
            ' --> move to the end after data fill in
        
        
        '------------------------------------------------
            ' Loop through all <tickers> rows and output
            '  -- Ticker Symbol
            '  -- Yearly Change
            '  -- Percent Change
            '  -- Total stock volume
        '-------------------------------------------------
        
        
        'set the first open price value
            openPrice = ws.Cells(2, 3)
        
        
        For row_num = 2 To lastRow
        
        ' check if we are still within the same ticker name
            ' if it's not:
            
             If ws.Cells(row_num + 1, 1).Value <> ws.Cells(row_num, 1).Value Then
             
                ' set the ticker
                ticker = ws.Cells(row_num, 1).Value
                
                ' add ticker name to the Summary Table
                ws.Range("I" & summary_table_row).Value = ticker
                
                ' counting a new total per ticker
                tickerTotal = tickerTotal + ws.Cells(row_num, 7).Value
                
                ' add the new total volume to the Summary Table
                ws.Range("L" & summary_table_row).Value = tickerTotal
                
        
                ' set the close price per ticker
                closePrice = ws.Cells(row_num, 6).Value
                
        
                ' counting yearly price change
                yearlyChange = closePrice - openPrice
               
                ' add to the Summary Table
                ws.Range("J" & summary_table_row).Value = yearlyChange
     
                    '----------------------------
                    '   conditional formatting
                    '----------------------------
                    ' apply conditional formatting to 'Yearly Change' column to highlight
                    ' -- positive change in Green (4)
                    ' -- negative change in Red (3)
                    
                    If yearlyChange >= 0 Then
                    
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    
                    Else
                    
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    
                    End If
                    
                 
                ' counting percent change
                
                    ' set format of Percent Change column in Summary Table to %
                    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                    
                    ' check condition for open price
                    If openPrice = 0 Then
    
                        ws.Range("K" & summary_table_row).Value = "0"
                    
                    Else
                        percentChange = yearlyChange / openPrice
                        
                        ws.Range("K" & summary_table_row).Value = percentChange
                    
                    End If
    
    
                ' move to the next row in the Summary Table
                summary_table_row = summary_table_row + 1
                
                
                ' set the next open price
                openPrice = ws.Cells(row_num + 1, 3).Value
                    
                              
                ' reset Total Volume per ticker
                tickerTotal = 0
                    
                    
                ' If the cell immediately following the current row has the same name
            Else
                
                ' counting a new total per ticker for the symbol already in the Summary Table
                tickerTotal = tickerTotal + ws.Cells(row_num, 7).Value
                
            End If
            
        Next row_num
        
        '------------------
        ' Bonus Part
        '------------------
        
        Dim sumTbl_lastRow As Double    ' determine last row in Summary Table
        Dim sumTbl_row_num As Double    ' counter for Summary Table rows

   
        '----------------------------------------------------------------------
        ' Analyze resulting data in the Summary Table to find out stocks with:
        '   -- the Greatest % Increase
        '   -- the Greatest % Decrease
        '   -- the Greatest Total Volume
        '----------------------------------------------------------------------


        sumTbl_lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
        
        For sumTbl_row_num = 2 To sumTbl_lastRow
        
          ' check if the value of current cell in 'Percent Change' column of Summary Table
            ' a) is maximum
            If ws.Cells(sumTbl_row_num, "K").Value = WorksheetFunction.Max(ws.Range("K2:K" & sumTbl_lastRow)) Then
                
                ' enter the ticker symbol for corresponding value on the 'Greatest % Increase' line of column 'Ticker' ('P')
                ws.Cells(2, "P") = ws.Cells(sumTbl_row_num, "I").Value
                                                              
                ' change formatting to %
                ws.Cells(2, "Q").NumberFormat = "0.00%"
                
                ' enter Percent Change maximum value on the 'Greatest % Increase' line of column 'Value' ('Q')
                ws.Cells(2, "Q").Value = WorksheetFunction.Max(ws.Range("K2:K" & sumTbl_lastRow))
                
                
            ' b) is minimum
            ElseIf ws.Cells(sumTbl_row_num, "K").Value = WorksheetFunction.Min(ws.Range("K2:K" & sumTbl_lastRow)) Then
            
                ' enter the ticker symbol for corresponding value on the 'Greatest % Decrease' line of column 'Ticker' ('P')
                ws.Cells(3, "P") = ws.Cells(sumTbl_row_num, "I").Value
                
                ' change formatting to %
                ws.Cells(3, "Q").NumberFormat = "0.00%"
                
                ' enter Percent Change minimum value on the 'Greatest % Decrease' line of column 'Value' ('Q')
                ws.Cells(3, "Q").Value = WorksheetFunction.Min(ws.Range("K2:K" & sumTbl_lastRow))

            ' if not a) or b)
            ' check if value of current cell in 'Total Stock Volume' column of Summary Table is maximum
            ElseIf ws.Cells(sumTbl_row_num, "L").Value = WorksheetFunction.Max(ws.Range("L2:L" & sumTbl_lastRow)) Then
            
                ' enter the ticker symbol for corresponding value on the 'Greatest Total Volume' line of column 'Ticker' ('P')
                ws.Cells(4, "P") = ws.Cells(sumTbl_row_num, "I").Value
                
                ' change formatting to "scientific"
                ws.Cells(4, "Q").NumberFormat = "0.0000E+00"
                
                ' enter Total Volume maximum value on the 'Greatest Total Volume' line of column 'Value' ('Q')
                ws.Cells(4, "Q").Value = WorksheetFunction.Max(ws.Range("L2:L" & sumTbl_lastRow))
                        
            End If
                
        Next sumTbl_row_num
        
        '---------------------------------------
        ' change columns width by autofit format
        '---------------------------------------
        
        ' add autofit column width formatting to the columns of Summary Tables
        ws.Range("I:Q").Columns.AutoFit
        
        
        '----------------------------------
        ' *-- use for testing purposes only
        '----------------------------------

'        ' clear all contents and formatting for added columns
'        ws.Range("I:Q").ClearContents
'        ws.Range("I:Q").ClearFormats
              


    Next ws

End Sub

    

