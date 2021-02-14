Sub CurrentSheet()

'PART 1 - Populate Summary Table

 Dim last_row As Long
 Dim ticker As String
 Dim table_row As Integer
 Dim open_price As Double
 Dim open_price2 As Double
 Dim closed_price As Double
 Dim yr_change As Double
 Dim total_vol As Double
 
 'Identify last row with data
 last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
 'Sort data with header by ticker and date (multi columns)
 ActiveSheet.Sort.SortFields.Clear
 ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 ActiveSheet.Sort.SortFields.Add Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 With ActiveSheet.Sort
 .SetRange Range("A1:G" & last_row)
 .Header = xlYes
 .MatchCase = False
 .Orientation = xlTopToBottom
 .Apply
 End With
    
 'Insert headers to analysis table
 Cells(1, 9) = "Ticker"
 Cells(1, 10) = "Yearly Change"
 Cells(1, 11) = "Percent Change"
 Cells(1, 12) = "Total Stock Volume"

 'Default initial values
 table_row = 2
 total_vol = 0
 open_price = Cells(2, 3).Value
 
 'Open price for zero division
 If open_price = 0 Then
  open_price2 = 1
 Else
  open_price2 = open_price
 End If
 
 For i = 2 To last_row
 
    'Check for change of ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
     'Unique ticker where a change occurs
     ticker = Cells(i, 1).Value
     
     'Add last row of each ticker to total stock volume
     total_vol = total_vol + Cells(i, 7).Value

     'Closed price
     closed_price = Cells(i, 6).Value
    
     'Print values to table
     Range("I" & table_row).Value = ticker
     Range("J" & table_row).Value = closed_price - open_price
     Range("K" & table_row).Value = FormatPercent(Round(((closed_price - open_price) / open_price2), 4))
     Range("L" & table_row).Value = total_vol
     
     'Colour formatting
     If Range("J" & table_row).Value > 0 Then
      Range("J" & table_row).Interior.Color = vbGreen
     ElseIf Range("J" & table_row).Value < 0 Then
      Range("J" & table_row).Interior.Color = vbRed
     Else
      Range("J" & table_row).Interior.ColorIndex = 34
      
     End If
     
     'Print for testing
     'Range("N" & table_row).Value = closed_price
     'Range("M" & table_row).Value = open_price
        
     'Next table_row
     table_row = table_row + 1
    
     'Reset total_vol
     total_vol = 0
     
     'Open price for next ticker
      open_price = Cells(i + 1, 3).Value
      
      'Open price for zero division
      If open_price = 0 Then
       open_price2 = 1
      Else
       open_price2 = open_price
      End If

    
    'If the cell immediately following a row is the same ticker
    Else
    
     'Running total stock volume
     total_vol = total_vol + Cells(i, 7).Value
     
    End If
 
 Next i


'PART 2 - Determine Greatest Values

 Dim last_tblrow As Long
 Dim maxp_ticker As String
 Dim minp_ticker As String
 Dim maxv_ticker As String
 Dim max_percent As Double
 Dim min_percent As Double
 Dim max_volume As Double
   
 'Identify last row with data
 last_tblrow = Cells(Rows.Count, 9).End(xlUp).Row
 
 'Max/min values
 max_percent = WorksheetFunction.Max(Range("K1:K" & last_tblrow))
 min_percent = WorksheetFunction.Min(Range("K1:K" & last_tblrow))
 max_volume = WorksheetFunction.Max(Range("L1:L" & last_tblrow))

 'Print values to table
 Cells(1, 16) = "Ticker"
 Cells(1, 17) = "Value"
 Cells(2, 15) = "Greatest % Increase"
 Cells(3, 15) = "Greatest % Decrease"
 Cells(4, 15) = "Greatest Total Volume"
 Range("Q2").Value = FormatPercent(max_percent)
 Range("Q3").Value = FormatPercent(min_percent)
 Range("Q4").Value = max_volume
 
 'Determine ticker with max/min values
 
 For i = 2 To last_tblrow
 
    'Check for ticker with max_percent
    If Cells(i, 11).Value = max_percent Then
     maxp_ticker = Cells(i, 9).Value
    
    ElseIf Cells(i, 11).Value = min_percent Then
     minp_ticker = Cells(i, 9).Value
     
    ElseIf Cells(i, 12).Value = max_volume Then
     maxv_ticker = Cells(i, 9).Value
     
    End If
 
 Next i
 
 'Print values to table
 Range("P2").Value = maxp_ticker
 Range("P3").Value = minp_ticker
 Range("P4").Value = maxv_ticker
 
End Sub

