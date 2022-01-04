Attribute VB_Name = "Module3"
Sub Stock_Multiple_Sheets()

    For Each ws In Worksheets

        'Declare variables and set-up dimensions
        Dim ticker As String
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percent_change As LongLong
        Dim total_volume As Double
        Dim lastRow As Long
        Dim start_value As Double

        'Setting up the summary table
        Dim rowcount As Integer
        rowcount = 2
        ticker_count = 1
        
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Add labels to summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        start_value = 2
   'Loop through all tickers
        For i = 2 To lastRow
            
            'Check if we are still within the same ticker name, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Assign values for the closing price
                 total_volume = total_volume + ws.Cells(i, 7).Value
                 
                 If ws.Cells(start_value, 3) = 0 Then
                    For find_nonzero_value = start_value To i
                        If ws.Cells(find_nonzero_value, 3).Value <> 0 Then
                            start_value = find_nonzero_value
                            
                            Exit For
                            
                        End If
                    Next find_nonzero_value
                
                End If
                
              
                yearly_change = ws.Cells(i, 6).Value - ws.Cells(start_value, 3).Value
                percent_change = Round((yearly_change / ws.Cells(start_value, 3).Value) * 100, 2)
                start_value = i + 1
                
                             
                'Populate values
                ws.Range("I" & rowcount).Value = ws.Cells(i, 1).Value
                ws.Range("J" & rowcount).Value = yearly_change
                ws.Range("K" & rowcount).Value = percent_change
                ws.Range("L" & rowcount).Value = total_volume
                
                'Formatting
                ws.Range("J" & rowcount).NumberFormat = "$#.##"
                ws.Range("K" & rowcount).NumberFormat = "0.00%"

                'Conditional formatting
                If ws.Range("J" & rowcount).Value < 0 Then
                    ws.Range("J" & rowcount).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & rowcount).Interior.ColorIndex = 10
                End If
                
                rowcount = rowcount + 1
                
            
            'Check if we are still within the same ticker name, if it is...
            Else
                
                'Assign values
                total_volume = total_volume + ws.Cells(i, 7).Value

                                       
                'Reset counters
                ticker_count = rowcount + 1
            
            End If
  
        Next i

    Next ws


End Sub

