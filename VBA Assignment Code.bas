Attribute VB_Name = "Module1"
Sub alphabetical_testing()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    Dim ticker_name As String
    Dim ticker_total As Double
    ticker_total = 0
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim summary_table_row As Integer
    summary_table_row = 2
    Dim first_row As Long
    first_row = 2
    last_row1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To last_row1
        
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
            ticker_name = ws.Cells(i, 1).Value
            yearly_change = ws.Cells(i, 6).Value - ws.Cells(first_row, 3).Value
            percentage_change = (ws.Cells(i, 6).Value - ws.Cells(first_row, 3).Value) / ws.Cells(first_row, 3).Value
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            ws.range("I" & summary_table_row).Value = ticker_name
            ws.range("L" & summary_table_row).Value = ticker_total
            ws.range("J" & summary_table_row).Value = yearly_change
            ws.range("K" & summary_table_row).Value = percentage_change
            ws.range("K" & summary_table_row).NumberFormat = "0.00%"
            summary_table_row = summary_table_row + 1
            first_row = i + 1
            ticker_total = 0
            
        Else
        
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
        End If
        

    Next i
    
    last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As LongLong
    Dim ticker_name1 As String
    Dim ticker_name2 As String
    Dim ticker_name3 As String
    
    max_increase = 0
    max_decrease = 0
    max_volume = 0
    
    For i = 2 To last_row2
        
        If (max_increase < ws.Cells(i, 11).Value) Then
            
            max_increase = ws.Cells(i, 11).Value
            ticker_name1 = ws.Cells(i, 9).Value
            
        Else
        
            i = i + 1
            
        End If
        
    Next i
    
    For i = 2 To last_row2
    
        If (max_decrease > ws.Cells(i, 11).Value) Then
        
            max_decrease = ws.Cells(i, 11).Value
            ticker_name2 = ws.Cells(i, 9).Value
            
        Else
        
            i = i + 1
            
        End If
        
    Next i
    
    For i = 2 To last_row2
    
        If (max_volume < ws.Cells(i, 12).Value) Then
        
            max_volume = ws.Cells(i, 12).Value
            ticker_name3 = ws.Cells(i, 9).Value
            
        Else
        
            i = i + 1
            
        End If
        
        ws.range("P2") = max_increase
        ws.range("P2").NumberFormat = "0.00%"
        ws.range("O2") = ticker_name1
        ws.range("P3") = max_decrease
        ws.range("P3").NumberFormat = "0.00%"
        ws.range("O3") = ticker_name2
        ws.range("P4") = max_volume
        ws.range("P4").NumberFormat = "0"
        ws.range("O4") = ticker_name3
        
    Next i
         
    For i = 2 To last_row2
    
        If (ws.Cells(i, 10).Value < 0) And (ws.Cells(i, 11).Value < 0) Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
            
        Else
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        End If
        
    Next i
    
Next
        
End Sub
