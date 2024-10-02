Attribute VB_Name = "Module1"

Sub Stock_Data()

'I came back and added a for loop and added "ws" behind each running formula and value to make the script
'run through all worksheets
For Each ws In Worksheets

'Setting the variables
Dim ticker As String
ticker = 0
Dim tickervolume As Double
tickervolume = 0
Dim summarytablerow As Integer
summarytablerow = 2
Dim openprice As Double
openprice = ws.Cells(2, 3)
Dim closeprice As Double
Dim quarterlychange As Double
Dim percentchange As Double


'Naming the Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Used this formula to subtract the closing value and the opening value to get the quarterly change
    closeprice = ws.Cells(i, 6).Value
    quarterlychange = (closeprice - openprice)
    ws.Range("J" & summarytablerow).Value = quarterlychange
    
    'Setting the formula for percentage (I had to use help for this one as the formula was not giving the division value,
    'https://stackoverflow.com/questions/62471422/vba-loop-how-to-get-ticker-symbols-into-ticker-column)
    If openprice = 0 Then
    percentchange = 0
    Else: percentchange = (quarterlychange / openprice)
    End If
    
    'Printing values on proper columns with percent sign
    'Had to google how to get percentage numbers, if I did this (quarterlychange / openprice) * 100
    'I would get values that were wrong
    'https://stackoverflow.com/questions/57814868/how-to-divide-and-get-the-percentage-of-
    'the-value-of-two-different-textboxes-in
    ws.Range("K" & summarytablerow).Value = percentchange
    ws.Range("K" & summarytablerow).NumberFormat = "0.00%"
        
    'I used an example from Module 2.3 with the "credit_charges solutions"
    ticker = ws.Cells(i, 1).Value
    tickervolume = tickervolume + ws.Cells(i, 7).Value
    ws.Range("I" & summarytablerow).Value = ticker
    ws.Range("L" & summarytablerow).Value = tickervolume
    
   'Reset the summarytablerow counter
    summarytablerow = summarytablerow + 1
    
   'Reset the voume to 0
    tickervolume = 0
    
    'Reset the opening price
    openprice = ws.Cells(i + 1, 3)
    
    Else
    'Adding the values of each ticker per ticker
    tickervolume = tickervolume + ws.Cells(i, 7).Value
 
End If
Next i

'Finding the last row of the summary table after getting results
'I set a value to find the last row of the new summary table with all the results
'This will be used later to get the max and min percentage increase
'As well as the greatest volume total volume
lastrowsummarytable = ws.Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To lastrowsummarytable

'Setting colors according to the percentages
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 10
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(i, 10).Value = 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 2
        
    End If
Next i

'Labeling the columns for the new requested values
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

For i = 2 To lastrowsummarytable

'Used google once again to get the function for max, then applied to the rest
'https://stackoverflow.com/questions/31906571/excel-vba-find-maximum-value-in-range-on-specific-sheet
If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowsummarytable)) Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowsummarytable)) Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
   ws.Cells(3, 17).NumberFormat = "0.00%"
    
ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowsummarytable)) Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    
End If
Next i
Next ws
End Sub
