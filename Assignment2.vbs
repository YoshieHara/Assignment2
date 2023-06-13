Sub Ticker()

For Each ws In Worksheets

Dim Ticker As String
Dim Total_Stock_Volume, Yearly_Change, Percent_Change, Year_Close, Year_Open As Double
Dim Summary_Table_Row As Integer
Dim Greatest_Increase As Double
Dim Greatest_decrease As Double
Dim Greatest_Total_Volume As Double
Summary_Table_Row = 2
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Total_Stock_Volume = 0
Year_Open = ws.Cells(2, 3).Value
Greatest_Increase = 0
Greatest_decrease = 0
Greatest_Total_Volume = 0

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

For i = 2 To Lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(Summary_Table_Row, 9) = Ticker
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ws.Cells(Summary_Table_Row, 12) = Total_Stock_Volume

      ' if the ticker is different from next row:
             'This row is the last one for this ticker
             'add this volume to total for the last time
             'go to the next row and set back the total to 0

        Year_Close = ws.Cells(i, 6).Value
        Yearly_Change = Year_Close - Year_Open
        ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
     
     'Conditional formatting
      If Yearly_Change < 0 Then
        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        
      ElseIf Yearly_Change > 0 Then
        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    
       End If
        
        
        Percent_Change = ((Year_Close / Year_Open) - 1)
        ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
        ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"



        If Percent_Change >= Greatest_Increase Then
            Greatest_Increase = Percent_Change
            ws.Cells(2, 17).Value = Greatest_Increase
            ws.Cells(2, 16).Value = Ticker
   
        End If
        
        If Percent_Change < Greatest_decrease Then
            Greatest_decrease = Percent_Change
            ws.Cells(3, 17).Value = Greatest_decrease
            ws.Cells(3, 16).Value = Ticker

               
        End If
        
        If Total_Stock_Volume > Greatest_Total_Volume Then
           Greatest_Total_Volume = Total_Stock_Volume
           ws.Cells(4, 17).Value = Greatest_Total_Volume
           ws.Cells(4, 16).Value = Ticker
        
        End If
        

        Year_Open = ws.Cells(i + 1, 3).Value
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Stock_Volume = 0


    Else
    'if ticker is the same as the next row, add the volume and run the loop again
     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value


    End If
Next i


Next ws

End Sub
