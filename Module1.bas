Attribute VB_Name = "Module1"
Sub MultipleYearStockData():
For Each ws In Worksheets

'Declaring Variables
Dim Worksheet_Name As String
Dim ticker_name As String
Dim ticker_Count As Long
Dim lrowA As Long
Dim lrowI As Long
Dim i As Long
Dim j As Long
Dim Per_Change As Double
Dim Great_Inc As Double
Dim Great_Dec As Double
Dim Great_Vol As Double
Worksheet_Name = ws.Name
'Creating headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
ticker_Count = 2
j = 2
lrowA = Cells(Rows.Count, 1).End(xlUp).Row
'Looping through the rows
For i = 2 To lrowA
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 ws.Cells(ticker_Count, 9).Value = ws.Cells(i, 1).Value
'calculating yearly change
ws.Cells(ticker_Count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

'Applying conditional formatting
If ws.Cells(ticker_Count, 10).Value < 0 Then
ws.Cells(ticker_Count, 10).Interior.ColorIndex = 3

Else
              
ws.Cells(ticker_Count, 10).Interior.ColorIndex = 4

End If

'Total volume
ws.Cells(ticker_Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
'Incrementing it by 1
ticker_Count = ticker_Count + 1
j = i + 1

   End If

Next i

lrowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

'summary
Great_Vol = ws.Cells(2, 12).Value
Great_Inc = ws.Cells(2, 11).Value
Great_Dec = ws.Cells(2, 11).Value

For i = 2 To lrowI
' check Greatest Value
If ws.Cells(i, 12).Value > Great_Vol Then
   Great_Vol = ws.Cells(i, 12).Value
   ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

Else

Great_Vol = Great_Vol

End If

If ws.Cells(i, 11).Value > Great_Inc Then
Great_Inc = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
Else
                
Great_Inc = Great_Inc
                
End If
                
 'Check greatest decrease
  If ws.Cells(i, 11).Value < Great_Dec Then
  Great_Dec = ws.Cells(i, 11).Value
  ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
  Else
                
  Great_Dec = Great_Dec
                
  End If
                
  'summary results
  ws.Cells(2, 17).Value = Format(Great_Inc, "Percent")
  ws.Cells(3, 17).Value = Format(Great_Dec, "Percent")
  ws.Cells(4, 17).Value = Format(Great_Vol, "Scientific")
            
            Next i
            
 'Adjust column width
 Worksheets(Worksheet_Name).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub

