Attribute VB_Name = "Module1"
Sub stock_data()

  For Each ws In Worksheets
    'Define the variables
    Dim i As Long
    Dim j As Long
    Dim RowCountA As Long
    Dim TickCount As Long
    Dim Ticker As String
    Dim RowCountI As Long
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    Dim TotalVol As Double
    Dim Quarterly As Double
    Dim SummaryRow As Double
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxVolume As Double
    Dim TickerMaxIncrease As String
    Dim TickerMaxDecrease As String
    Dim TickerMaxVolume As String
    Dim RowCountFinal As Long
    
    
    
   
    
    'Add Summary Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Add headers for calculated values
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Volume"
    
    'Set the starting value for each variable
    Ticker = ""
    OpenPrice = 0
    ClosingPrice = 0
    Quarterly = 0
    TotalVol = 0
    SummaryRow = 2
    TickCount = 2
    j = 2
    
     'Add in a way to signify which row is the last row
    RowCountA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To RowCountA
            'Add in script that will run through each value in "ticker" and note when the ticker has changed
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Collect the total value of the Total Stock Volume
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                'Add the Ticker to my summary data
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                'Calculate the Quarterly Change using Closing Price - Open Price
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional Formatting for Quarterly Change, Green if the number is positive, Red if it is Negative
                    If ws.Cells(TickCount, 10) > 0 Then
                        ws.Cells(TickCount, 10).Interior.Color = RGB(0, 128, 0)
                    ElseIf ws.Cells(TickCount, 10).Value < 0 Then
                        ws.Cells(TickCount, 10).Interior.Color = RGB(255, 0, 0)
                    End If
                
                'Calculate the Percent Change if the Quarterly Change is not zero
                If ws.Cells(TickCount, 10).Value <> 0 Then
                    'Percent Change Calculation (Quarterly/Open Price)
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    'Format Percent Change
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    'Conditional Formatting for Percent Change, Green if Positive, Red if Negative
                    If ws.Cells(TickCount, 11) > 0 Then
                        ws.Cells(TickCount, 11).Interior.Color = RGB(0, 128, 0)
                    ElseIf ws.Cells(TickCount, 11).Value < 0 Then
                        ws.Cells(TickCount, 11).Interior.Color = RGB(255, 0, 0)
                    End If
                'Set Percent Change to Zero if the Quarterly Change is Zero
                Else
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                End If
                j = i + 1
                TickCount = TickCount + 1
            End If
      
      Next i
      Next ws
    
    'Initialize Variables to track largest values and associated Tickers
    MaxPercentIncrease = 0
    MaxPercentDecrease = 0
    MaxVolume = 0
    TickerMaxIncrease = ""
    TickerMaxDecrease = ""
    TickerMaxVolume = ""
    
    For Each ws In Worksheets
    RowCountFinal = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        For i = 2 To ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
            'Check for Largest % Increase
            If ws.Cells(i, 11).Value > MaxPercentIncrease Then
                MaxPercentIncrease = ws.Cells(i, 11).Value
                TickerMaxIncrease = ws.Cells(i, 9).Value
            End If
            'Check for largest % Decrease
            If ws.Cells(i, 11).Value < MaxPercentDecrease Then
                MaxPercentDecrease = ws.Cells(i, 11).Value
                TickerMaxDecrease = ws.Cells(i, 9).Value
            End If
            'Check for largest total volume
            If ws.Cells(i, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(i, 12).Value
                TickerMaxVolume = ws.Cells(i, 9).Value
            End If
        Next i
        
        'Output the results to summary table
        ws.Cells(2, 16).Value = MaxPercentIncrease
            'Format Percent Change
            ws.Cells(2, 16).Value = Format(MaxPercentIncrease, "Percent")
        ws.Cells(3, 16).Value = MaxPercentDecrease
            'Format Percent Change
            ws.Cells(3, 16).Value = Format(MaxPercentDecrease, "Percent")
        ws.Cells(4, 16).Value = MaxVolume
        ws.Cells(2, 15).Value = TickerMaxIncrease
        ws.Cells(3, 15).Value = TickerMaxDecrease
        ws.Cells(4, 15).Value = TickerMaxVolume
    Next ws
    
End Sub
 
           
           
           
           
           
