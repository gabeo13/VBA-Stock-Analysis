Attribute VB_Name = "VBAstockanalysis"
Sub stockanalysis():

'1. Discern & print unique/distinct ticker symbols
'2. Summation of total stock volume by symbol
'3. Calculate and print yearly change by symbol
'4. Format positive yearly change green and negative red
'5. Calculate and print percent yearly change by symbol
'6. Discern and print stock with greatest % increase
'7. Discern and print stock with greatest % decrease
'8. Discern and print stock with most volume traded
'9. Loop through each worksheet in workbook

'Declare Variables
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim output_row As Long
Dim total_vol As Variant
Dim last_row As Long
Dim ws As Worksheet


    'Iterate through each sheet in workbook (outermost loop)
    For Each ws In Worksheets
    
    'Print Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Largest Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N1").Value = "Bonus Material"
    
    'Start Counter and define start of output
    total_vol = 0
    output_row = 2
    
    'Define last row to bind the for loop range
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Inner for loop to iterate through data set
        For i = 2 To last_row
        
            'Outer conditional to extract discrete and calculated values
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set distinct ticker value to that of last known occurance
                ticker = ws.Cells(i, 1).Value
                
                'Summate final total volume value with running total for the ith symbol
                total_vol = total_vol + ws.Cells(i, 7).Value
                
                'Extract discrete values for close and open price from last and first row
                close_price = ws.Cells(i, 6).Value
                open_price = ws.Cells(i - (i - 2), 3).Value
                
                'Calculate yearly change and percent change
                yearly_change = close_price - open_price
                percent_change = yearly_change / open_price
                
                'Push values to summary table output row
                ws.Range("I" & output_row).Value = ticker
                ws.Range("L" & output_row).Value = total_vol
                ws.Range("J" & output_row).Value = yearly_change
                ws.Range("k" & output_row).Value = percent_change
                
                    'Inner conditional to colorize yearly change
                    If yearly_change > 0 Then
                        ws.Range("J" & output_row).Interior.ColorIndex = 43
                    ElseIf yearly_change < 0 Then
                        ws.Range("J" & output_row).Interior.ColorIndex = 3
                    Else
                    End If
                    
                'Increment output row and reset total volume counter
                output_row = output_row + 1
                total_vol = 0
            Else
                'Summate total volume up to last row of each distinct ticker value
                total_vol = total_vol + ws.Cells(i, 7).Value
            End If
            
        'Cell iterator --> push to next
        Next i
        
    'Extract requested values from output table
    ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("P3").Value = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L:L"))
    
    'Match requested values from above to their ticker in column "I"
    ws.Range("O2").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("P2").Value, ws.Range("K:K"), 0))
    ws.Range("O3").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("P3").Value, ws.Range("K:K"), 0))
    ws.Range("O4").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("P4").Value, ws.Range("L:L"), 0))
    
    'Column & formatting clean up
    ws.UsedRange.Columns.AutoFit
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("P2", "P3").NumberFormat = "0.00%"
             
    'Worksheet iterator --> push to next
    Next ws
    
'close sub procedure
End Sub
