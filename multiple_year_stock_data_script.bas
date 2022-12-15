Attribute VB_Name = "Module1"
'----------------------------------
'Create a script that loops through all the stocks for one year and outputs the following information:
    'The ticker symbol
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
'----------------------------------

Sub stock_market()

'Repeat for all worksheets in this workbook
Dim ws As Worksheet
For Each ws In Worksheets

'Set column headers for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'Define variables for all loops
    Dim ticker_LastRow As Long
    Dim first_date As Long
    Dim last_date As Long
    Dim first_date_opening As Double
    Dim last_date_closing As Double
    
    
    ticker_row = 2 'start writing down values in the ticker col from row 2
    yearly_row = 2 'start writing down values in yearly change col from row 2
    percent_row = 2 'start writing down values in percent change col from row 2
    total_row = 2 'start writing down values in total stock volume col from row 2
    ticker_LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'get total number of rows
    
'----------------------------------
    'Loop through each sheet to get ticker value
    
   
        For i = 2 To ticker_LastRow
        
        total_volume = ws.Cells(i, 7).Value + total_volume 'start adding up the stock volumes
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then 'if the loop reaches the last row for each unique ticker
                ws.Cells(ticker_row, 9).Value = ws.Cells(i, 1).Value 'copy the ticker value to Ticker column
                ticker_row = ticker_row + 1 'move down row in ticker column

                last_date_closing = ws.Cells(i, 6).Value 'loop through until it reaches the last row
                ws.Cells(yearly_row, 10).Value = last_date_closing - first_date_opening 'once last row is reached, calculate the yearly change and copy to summary column
                    If Sgn(ws.Cells(yearly_row, 10).Value) = 1 Then
                        ws.Cells(yearly_row, 10).Interior.ColorIndex = 4
                    ElseIf Sgn(ws.Cells(yearly_row, 10).Value) = -1 Then
                        ws.Cells(yearly_row, 10).Interior.ColorIndex = 3
                    End If
                yearly_row = yearly_row + 1 'in yearly row column, go to next row
                
                ws.Cells(percent_row, 11).Value = ((ws.Cells(percent_row, 10).Value) / first_date_opening) * 100 & "%" 'calculate the percent change and write it down
                percent_row = percent_row + 1 'move to the next row in percent change col
                
                ws.Cells(total_row, 12).Value = total_volume 'write down the total stock volumn in the corresponding cell
                total_row = total_row + 1 'move to the next row in total stock volumn col
                
                total_volume = 0 'set total volume to 0 on last row
                
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then 'if it is the first row
                    first_date_opening = ws.Cells(i, 3).Value 'then index the first row and store the price
            End If
        Next i
        
Next ws

End Sub

