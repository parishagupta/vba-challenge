Attribute VB_Name = "Module1"
Sub VBAStocksChallengeMain():

'Step 1 For Loop WS
Dim ws As Worksheet
For Each ws In Worksheets

'Step 2 Define All Variables
Dim Ticker As String
Dim Total_Stock_Volume As Variant
    'Set to 0 to start
    Total_Stock_Volume = 0
Dim Summary_Table_Row As Integer
    'To keep track of location for each Ticker in summary table starting in row 2
    Summary_Table_Row = 2
Dim Quarter_Open As Variant
Dim Quarter_Close As Variant
Dim Quarterly_Change As Variant
Dim Percentage_Change As Variant
Dim Greatest_increase As Variant
    Greatest_increase = 0
Dim Greatest_decrease As Variant
    Greatest_decrease = 0
Dim Greatest_Volume As Variant
    Greatest_Volume = 0
'Define last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Step 3 Display additional columns headers
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Quarterly Change"
ws.Cells(1, 12).Value = "Percentage Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
    
'Step 4 For Loop - loop through all data rows and display in step 3 columns
For i = 2 To lastrow
    
    'Step 4.2 If current and next cells are not equal then display the i string of text and totals in columns 10 and 13
    If ws.Cells(1 + i, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ws.Range("J" & Summary_Table_Row).Value = Ticker
        ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
        'Formula and display Quarterly Change $0.00 in column 11
        Quarter_Close = ws.Cells(i, 6).Value
        Quarter_Open = ws.Cells(i - 61, 3).Value
        Quarterly_Change = (Quarter_Close - Quarter_Open)
        ws.Range("K" & Summary_Table_Row).Value = Quarterly_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "$0.00"
        'Formula and display Percentage Change 0.00% in column 12
        Percentage_Change = (Quarterly_Change / Quarter_Open)
        ws.Range("L" & Summary_Table_Row).Value = Percentage_Change
        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Step 4.3 color conditional formatting for Quarterly Change and Percentage Change
                If ws.Range("K" & Summary_Table_Row).Value >= 0 Then
                     ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
                     ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                If ws.Range("L" & Summary_Table_Row).Value >= 0 Then
                     ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf ws.Range("L" & Summary_Table_Row).Value < 0 Then
                     ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
        
        'Step 4.4 add 1 row to Summary Table
        Summary_Table_Row = (Summary_Table_Row + 1)

        
        'Step 4.5 display the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
            If ws.Cells(i, 12).Value > Greatest_increase Then
                Greatest_increase = ws.Cells(i, 12).Value
                ws.Cells(2, 18).Value = Greatest_increase
                ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
                ws.Cells(2, 18).NumberFormat = "0.00%"
            ElseIf ws.Cells(i, 12).Value < Greatest_decrease Then
                Greatest_decrease = ws.Cells(i, 12).Value
                ws.Cells(3, 18).Value = Greatest_decrease
                ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
                ws.Cells(3, 18).NumberFormat = "0.00%"
            ElseIf ws.Cells(i, 13).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(i, 13).Value
                ws.Cells(4, 18).Value = Greatest_Volume
                ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
            End If
            
            ws.Columns("J:R").AutoFit
                                
     'If the next row is the same Ticker then add to Volume
    Else
        Quarter_Open = ws.Cells(i, 3).Value
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    End If
                      
        'Reset Variables
        Total_Stock_Volume = 0
        Quarter_Open = 0
        Quarter_Close = 0
        Percentage_Change = 0
        
        'Reset Variables
        Greatest_increase = 0
        Greatest_decrease = 0
        Greatest_Volume = 0

Next i
 
Next ws

End Sub

