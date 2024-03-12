Attribute VB_Name = "Module1"
Sub Check_Data()

'Initial Data compilation

    ' Declare Variables
    Dim ws As Worksheet
    Dim ticker_pos As Integer
    Dim ticker_name As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim last_row As Long
    Dim year_start As Double
    Dim year_end As Double
    
    
    
    
    'For loop cycles worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    
    'Set initial numeric values for sheet
    ticker_pos = 2
    ticker_name = ""
    yearly_change = 0
    percent_change = 0
    total_volume = 0
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    year_start = ws.Cells(2, 3).Value
    
    
    'Begin iteration to pull ticker names
    For i = 2 To last_row
        If ticker_name = "" Then
            ticker_name = ws.Cells(i, 1).Value
       End If
            
       'Sum total volume per ticker
       total_volume = total_volume + ws.Cells(i, 7).Value
       
       
       'If statement ends data collection for each ticker
       If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then      '???only works if ticker names and dates are ordered???
            
            
            'Set/Calculate Values; end of year, yearly change, percent change
            year_end = ws.Cells(i, 6).Value
            yearly_change = year_end - year_start
            percent_change = yearly_change / year_start
            
            
            'Print ticker values; name, yearly change, percent change, total volume
            ws.Cells(ticker_pos, 9).Value = ticker_name
            ws.Cells(ticker_pos, 10).Value = yearly_change              '=
            ws.Cells(ticker_pos, 11).Value = percent_change
            ws.Cells(ticker_pos, 12).Value = total_volume                 '
            
            
            'Reset variables; year start, ticker name, total volume, ticker position
            year_start = ws.Cells(i + 1, 3).Value
            ticker_name = ws.Cells(i + 1, 1).Value
            total_volume = 0
            ticker_pos = ticker_pos + 1
        End If
        
    Next i
            
            
            
'For loop for table of greatest values.
   
   
   'Declare variables
   Dim last_table_row As Long
   Dim greatest_percent_value As Double
   Dim greatest_percent_name As String
   Dim lowest_percent_value As Double
   Dim lowest_percent_name As String
   Dim greatest_volume_value As Double
   Dim greatest_volume_name As String
   
   
   'Set initial numeric values
    last_table_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    greatest_percent_value = 0
    lowest_percent_value = 0
    greatest_volume_value = 0
    
  
    'Begin iteration for greatest values
    For j = 2 To last_table_row
    
   
         'Find greatest percent increase
        If ws.Cells(j, 11).Value > greatest_percent_value Then
            greatest_percent_value = ws.Cells(j, 11).Value
            greatest_percent_name = ws.Cells(j, 9).Value
        End If
        
         'Find greatest percent decrease
        If ws.Cells(j, 11).Value < lowest_percent_value Then
            lowest_percent_value = ws.Cells(j, 11).Value
            lowest_percent_name = ws.Cells(j, 9).Value
        End If
        
         'Find greatest total volume
        If ws.Cells(j, 12).Value > greatest_volume_value Then
            greatest_volume_value = ws.Cells(j, 12).Value
            greatest_volume_name = ws.Cells(j, 9).Value
        End If
    
    Next j
    
    'Print Values
    ws.Cells(2, 16).Value = greatest_percent_name
    ws.Cells(3, 16).Value = lowest_percent_name
    ws.Cells(4, 16).Value = greatest_volume_name
    ws.Cells(2, 17).Value = greatest_percent_value
    ws.Cells(3, 17).Value = lowest_percent_value
    ws.Cells(4, 17).Value = greatest_volume_value

 Next ws
 
End Sub
