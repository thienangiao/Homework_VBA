Attribute VB_Name = "Multisheet_Update"
Sub MultiSheets()

 Dim last_row As Long
 Dim ticker As String
 Dim table_row As Integer
 Dim open_price As Double
 Dim open_price2 As Double
 Dim closed_price As Double
 Dim yr_change As Double
 Dim total_vol As Double
 
 Dim last_tblrow As Long
 Dim maxp_ticker As String
 Dim minp_ticker As String
 Dim maxv_ticker As String
 Dim max_percent As Double
 Dim min_percent As Double
 Dim max_volume As Double
  
 'Loop through all sheets
 For Each ws In Worksheets
 
     'PART 1 - Populate Summary Table
     
     'Identify last row with data
     last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'Sort data with header by ticker and date (multi columns)
     ws.Sort.SortFields.Clear
     ws.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     ws.Sort.SortFields.Add Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     With ws.Sort
     .SetRange Range("A1:G" & last_row)
     .Header = xlYes
     .MatchCase = False
     .Orientation = xlTopToBottom
     .Apply
     End With
        
     'Insert headers to analysis table
     ws.Cells(1, 9) = "Ticker"
     ws.Cells(1, 10) = "Yearly Change"
     ws.Cells(1, 11) = "Percent Change"
     ws.Cells(1, 12) = "Total Stock Volume"
        
     'Default initial values
     table_row = 2
     total_vol = 0
     open_price = ws.Cells(2, 3).Value
     
    'Open price for zero division
     If open_price = 0 Then
      open_price2 = 1
     Else
      open_price2 = open_price
     End If
     
     For i = 2 To last_row
     
        'Check for change of ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
         'Unique ticker where a change occurs
         ticker = ws.Cells(i, 1).Value
         
         'Add last row of each ticker to total stock volume
         total_vol = total_vol + ws.Cells(i, 7).Value
    
         'Closed price
         closed_price = ws.Cells(i, 6).Value
        
         'Print values to table
         ws.Range("I" & table_row).Value = ticker
         ws.Range("J" & table_row).Value = closed_price - open_price
         ws.Range("K" & table_row).Value = FormatPercent(Round(((closed_price - open_price) / open_price2), 4))
         ws.Range("L" & table_row).Value = total_vol
         
         'Colour formatting
         If ws.Range("J" & table_row).Value > 0 Then
          ws.Range("J" & table_row).Interior.Color = vbGreen
         ElseIf ws.Range("J" & table_row).Value < 0 Then
          ws.Range("J" & table_row).Interior.Color = vbRed
         Else
          ws.Range("J" & table_row).Interior.ColorIndex = 34
          
         End If
         
         'Print for testing
         'ws.Range("N" & table_row).Value = closed_price
         'ws.Range("M" & table_row).Value = open_price
            
         'Next table_row
         table_row = table_row + 1
        
         'Reset total_vol
         total_vol = 0
         
         'Open price for next ticker
          open_price = ws.Cells(i + 1, 3).Value
          
          'Open price for zero division
          If open_price = 0 Then
           open_price2 = 1
          Else
           open_price2 = open_price
          End If
    
        
        'If the cell immediately following a row is the same ticker
        Else
        
         'Running total stock volume
         total_vol = total_vol + ws.Cells(i, 7).Value
         
        End If
     
     Next i
     
     'PART 2 - Determine Greatest Values
     
     'Identify last row with data
     last_tblrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
     'Max/min values
     max_percent = WorksheetFunction.Max(ws.Range("K1:K" & last_tblrow))
     min_percent = WorksheetFunction.Min(ws.Range("K1:K" & last_tblrow))
     max_volume = WorksheetFunction.Max(ws.Range("L1:L" & last_tblrow))
    
     'Print values to table
     ws.Cells(1, 16) = "Ticker"
     ws.Cells(1, 17) = "Value"
     ws.Cells(2, 15) = "Greatest % Increase"
     ws.Cells(3, 15) = "Greatest % Decrease"
     ws.Cells(4, 15) = "Greatest Total Volume"
     ws.Range("Q2").Value = FormatPercent(max_percent)
     ws.Range("Q3").Value = FormatPercent(min_percent)
     ws.Range("Q4").Value = max_volume
     
     'Determine ticker with max/min values
     
     For i = 2 To last_tblrow
     
        'Check for ticker with max_percent
        If ws.Cells(i, 11).Value = max_percent Then
         maxp_ticker = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 11).Value = min_percent Then
         minp_ticker = ws.Cells(i, 9).Value
         
        ElseIf ws.Cells(i, 12).Value = max_volume Then
         maxv_ticker = ws.Cells(i, 9).Value
         
        End If
     
     Next i
     
     'Print values to table
     ws.Range("P2").Value = maxp_ticker
     ws.Range("P3").Value = minp_ticker
     ws.Range("P4").Value = maxv_ticker
 
  
 Next ws

End Sub

