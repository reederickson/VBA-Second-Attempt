# VBA-Second-Attempt
My excel file was too large to upload, even in a compressed form. Here is my second attempt at the code:
-side note, I got the correct results! After spending hours on stack overflow and rewatching the lecture, I found it was easier to restart my code from scratch rather than edit what I had previously attempted. The overall code looks a lot prettier and best yet, actually runs.

Sub tickerloop()

'Loop through all worksheets
For Each ws In Worksheets

'ticker name
Dim tickername As String
    
'total count on the total volume
Dim tickervolume As Double
tickervolume = 0

'Keep track of ticker name in the summary table
Dim summary_ticker_row As Integer
summary_ticker_row = 2

Dim open_price As Double
'Set initial open_price
open_price = Cells(2, 3).value
        
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Summary Table headers
Cells(1, 9).value = "Ticker"
Cells(1, 10).value = "Yearly Change"
Cells(1, 11).value = "Percent Change"
Cells(1, 12).value = "Total Stock Volume"

'Count the number of rows in the first column.
LastRow = Cells(Rows.Count, 1).End(xlUp).row

'Loop through the rows by the ticker names
For i = 2 To LastRow

        'Searches for when the value of the next cell is different than that of the current cell
        If Cells(i + 1, 1).value <> Cells(i, 1).value Then

            'Set the ticker name
            tickername = Cells(i, 1).value

            'Add the volume
            tickervolume = tickervolume + Cells(i, 7).value

            'Print the ticker name in the summary table
            Range("I" & summary_ticker_row).value = tickername

            'Print the volume for each ticker in the summary table
            Range("L" & summary_ticker_row).value = tickervolume

            'closing price
            close_price = Cells(i, 6).value

            'Calculate yearly change
            yearly_change = (close_price - open_price)
              
            'Print yearly change for each ticker in the summary table
            Range("J" & summary_ticker_row).value = yearly_change

            'Check for the non-divisibilty
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / open_price
            End If

'Print yearly change for each ticker in the summary table
Range("K" & summary_ticker_row).value = percent_change
Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
'Reset the row counter. Add one to the summary_ticker_row
summary_ticker_row = summary_ticker_row + 1

'Reset volume
tickervolume = 0

'Reset opening price
open_price = Cells(i + 1, 3)
            
Else
              
'Add the volume of trade
tickervolume = tickervolume + Cells(i, 7).value

            
End If
        
        Next i

'Conditional formatting

lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i


'Set up summary table
Cells(2, 15).value = "Greatest % Increase"
Cells(3, 15).value = "Greatest % Decrease"
Cells(4, 15).value = "Greatest Total Volume"
Cells(1, 16).value = "Ticker"
Cells(1, 17).value = "Value"

'Determine the max and min values in column "Percent Change" and just max in column "Total Stock Volume"
'Collect the ticker name, with respective percent change and total volume
        For i = 2 To lastrow_summary_table
            'Maximum percent change
            If Cells(i, 11).value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).value = Cells(i, 9).value
                Cells(2, 17).value = Cells(i, 11).value
                Cells(2, 17).NumberFormat = "0.00%"

            'Minimum percent change
            ElseIf Cells(i, 11).value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).value = Cells(i, 9).value
                Cells(3, 17).value = Cells(i, 11).value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Maximum volume
            ElseIf Cells(i, 12).value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).value = Cells(i, 9).value
                Cells(4, 17).value = Cells(i, 12).value
            
             End If
        
         Next i
        
     Next ws
    
End Sub
