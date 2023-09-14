# Kevin-MONU-Module-2
vbahomework
Sub ticker_table()

'define everything
Dim ws As Worksheet
Dim ticker As String
Dim total_stock_volume As Long
total_stock_volume = 0
Dim yearly_open As Double
Dim yearly_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim ticker_table_row As Integer


'preventing overflow
On Error Resume Next - https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement This was retrieved because there were overflow and the values were unable to be obtained.
'run through each worksheet
For Each ws In ThisWorkbook.Worksheets - Code were established with the tutor, most of the ws parts were assisted "ws." pretty much applies to all current worksheets.
    'set headers 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup values for loop
    ticker_table_row = 2
    
    yearly_open = ws.Cells(2, 3) - Initially the code was actually within the loop, the instructor assisted with dragging the yearly_open out of the loop for to create allow the initial open price value.
    
    Last_Row = ws.Range("A" & ws.Rows.Count).End(xlUp).Row 

        'loop
        For i = 2 To Last_Row - Since we couldn't use static values as this loop needed to be made across multiple worksheets I had to define Last Row as a dynamic value to cater for the different values of rows in the different sheets.
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then - Previous codes used in the credit card exercise. When i,1 value differs form the next row
            
            'find all the values
            ticker = ws.Cells(i, 1)
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

            yearly_close = ws.Cells(i, 6).Value

            yearly_change = yearly_close - yearly_open
            
            percent_change = (yearly_close - yearly_open) / yearly_open
            
            'insert values into summary
            ws.Range("I" & ticker_table_row).Value = ticker
            ws.Range("J" & ticker_table_row).Value = yearly_change
            ws.Range("K" & ticker_table_row).Value = percent_change
            ws.Range("L" & ticker_table_row).Value = total_stock_volume
    
            ticker_table_row = ticker_table_row + 1

            total_stock_volume = 0
            
            yearly_open = ws.Cells(i + 1, 3).Value - Instructor assisted helped me through adding this statement, this uses the value next cell after the last cells open yearly open value.
            
            Else
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
        End If

    'finish loop
    Next i
    
'change the values of rows K to become percentage
ws.Columns("K").NumberFormat = "0.00%"
    'define variables for colours - https://www.easytweaks.com/excel-vba-change-cell-color/ Used this as a reference to assist me with the colour coding.
    Dim color_range As Range
    Dim j As Long
    Dim x As Long
    Dim color_cell As Range
    
    'setting values for loop
    Set color_range = ws.Range("J2", Range("J2").End(xlDown))
    x = color_range.Cells.Count
    
    'set colours for cells
    For j = 1 To x
    Set color_cell = color_range(j)
    Select Case color_cell
    'make conditional statements
        Case Is >= 0
            With color_cell
                .Interior.ColorIndex = 4
            End With
            
        Case Is < 0
            With color_cell
                .Interior.ColorIndex = 3
            End With
       End Select
    Next j


'move to next worksheet
Next ws


End Sub
