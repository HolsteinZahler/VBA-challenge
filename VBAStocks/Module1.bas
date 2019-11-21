Attribute VB_Name = "Module1"
Sub stocks():

Dim j As Long
Dim current_tick As String
Dim current_volume As LongLong
Dim current_index As Integer
Dim start_price As Double
Dim percent_change As Double
Dim change As Double
Dim greatest_inc As Double
Dim greatest_inc_tick As String
Dim greatest_dec As Double
Dim greatest_dec_tick As String
Dim greatest_volume As LongLong
Dim greatest_volume_tick As String
Dim lastrow As Long


lastrow = Cells(Rows.Count, 1).End(xlUp).Row

current_tick = Cells(2, 1).Value
current_volume = 0
current_index = 1
start_price = Cells(2, 3).Value

greatest_inc = 0   'Assuming at least one stock has positive return
greatest_dec = 0   'Assuming at least one stock has negative return
greatest_volume = 0

'Set table column labels
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"


For j = 2 To lastrow

    If Cells(j, 1) = current_tick Then
    
        current_volume = current_volume + Cells(j, 7)
        
    Else
        'Finalizing values for the current stock.
        'Cells(j - 1, 6) is the last closing cost
        change = Cells(j - 1, 6) - start_price
        'current_index is the row to be modified in the table to the
        'right of the data
        current_index = current_index + 1
        If start_price <> 0 Then
            percent_change = change / start_price
            Cells(current_index, 11) = percent_change
            Cells(current_index, 11).NumberFormat = "0.00%"
            If percent_change > greatest_inc Then
                greatest_inc = percent_change
                greatest_inc_tick = current_tick
            ElseIf percent_change < greatest_dec Then
                greatest_dec = percent_change
                greatest_dec_tick = current_tick
            End If
        Else
            Cells(current_index, 11) = "NaN"
        End If

        Cells(current_index, 9) = current_tick
        Cells(current_index, 10) = change
        If change > 0 Then 'Cell Color Green If Change is Positive
            Cells(current_index, 10).Interior.ColorIndex = 4
        ElseIf change < 0 Then 'Cell Color Red If Change is Negative
            Cells(current_index, 10).Interior.ColorIndex = 3
        Else 'Cell Color Yellow If Change is Zero
            Cells(current_index, 10).Interior.ColorIndex = 6
        End If
        
        If current_volume > greatest_volume Then
            greatest_volume = current_volume
            greatest_volume_tick = current_tick
        End If
        
        'Initialize values for next stock
        Cells(current_index, 12).Value = current_volume
        
        current_volume = Cells(j, 7).Value
        current_tick = Cells(j, 1).Value
        start_price = Cells(j, 3).Value
        

    
    End If
    

Next j

worksheet_name = ActiveSheet.Name

'Make Bonus Summary Table
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(2, 16) = greatest_inc_tick
Cells(2, 17) = greatest_inc
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 15) = "Greatest % Decrease"
Cells(3, 16) = greatest_dec_tick
Cells(3, 17) = greatest_dec
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 15) = "Greatest Total Volume"
Cells(4, 16) = greatest_inc_tick
Cells(4, 17) = greatest_volume

'Resize Columns
Columns("I:L").AutoFit
Columns("O:Q").AutoFit

'output_string = "Stock with greatest increase in " & worksheet_name & " was " & greatest_inc_tick
'MsgBox (output_string)

End Sub



