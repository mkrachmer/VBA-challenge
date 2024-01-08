Attribute VB_Name = "Module11"
Sub StockChallenge():

'Create a script that loops through all the stocks for one year and outputs the following information:

 ' Loop through all sheets
Dim ws As Worksheet
    
For Each ws In Worksheets

    'Create headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'ws.Cells(1, 13).Value = "OpenPrice" 'NOT NEEDED - JUST FOR TESTING
    'ws.Cells(1, 14).Value = "ClosePrice" 'NOT NEEDED - JUST FOR TESTING
    
    'Output the unique ticker symbols
    Dim Ticker As String
    
    'Counter for ticker symbols and initialize to 1
    Dim I As Long
    Dim NewTickerCounter As Integer
    NewTickerCounter = 1
    
    'Counter for the number of rows; referenced in day 3 module
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row 'xlUp is a lc L not a number 1
        'LastRow allows for differing row counts across tabs
        'MsgBox (LastRow) 'verify row count total
    
    'Variables for price calc
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YrlyChg As Double
    Dim PriceChgPct As Double
    Dim DailyVolume As Double
    Dim TotalVolume As LongLong
    Dim High As Double
    Dim Low As Double

    'OpenPrice = 0
    'ClosePrice = 0
    'YrlyChg = 0
    'PriceChgPct = 0
        
        'Print first ticker and open price
        Ticker = ws.Cells(2, 1).Value
          ws.Cells(2, 9).Value = Ticker
        OpenPrice = ws.Cells(2, 3).Value
          'ws.Cells(2, 13).Value = OpenPrice 'NOT NEEDED - JUST FOR TESTING
        
        NewTickerCounter = NewTickerCounter + 1
        
        'set daily volume to add to total volume in for loop
        DailyVolume = ws.Cells(2, 7).Value
        
        'Loop through all rows to pull new ticker, prices and volume
        For I = 2 To LastRow
          
          'check if ticker matches ticker on previous line. If it doesn't...
           If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
              'Add daily volume to the Volume Total, print, then reset the total to 0
             DailyVolume = ws.Cells(I, 7).Value
             TotalVolume = TotalVolume + DailyVolume
                 ws.Cells(NewTickerCounter, 12).Value = TotalVolume
             TotalVolume = 0
             'increment the ticket counter
             NewTickerCounter = NewTickerCounter + 1
             'Print the close price
             ClosePrice = ws.Cells(I, 6).Value
                 'ws.Cells(NewTickerCounter - 1, 14).Value = ClosePrice  'NOT NEEDED - JUST FOR TESTING
             'Print the next ticker
             Ticker = ws.Cells(I + 1, 1).Value
                 ws.Cells(NewTickerCounter, 9).Value = Ticker
             
             'Calculate and print the ticker yearly $ change
             YrlyChg = (ClosePrice - OpenPrice)
                 ws.Cells(NewTickerCounter - 1, 10).Value = YrlyChg
             'If positive change highlight in green. If negative highlight in red.
             If YrlyChg >= 0 Then
                ws.Cells(NewTickerCounter - 1, 10).Interior.ColorIndex = 4
             ElseIf YrlyChg < 0 Then
                ws.Cells(NewTickerCounter - 1, 10).Interior.ColorIndex = 3
             End If
             
             'Calculate and print ticker yearly % change
             PriceChgPct = (ClosePrice - OpenPrice) / OpenPrice
                 ws.Cells(NewTickerCounter - 1, 11).Value = FormatPercent(PriceChgPct)
             'If positive % change highlight in green. If negative highlight in red.
             If YrlyChg >= 0 Then
                ws.Cells(NewTickerCounter - 1, 11).Interior.ColorIndex = 4
             ElseIf YrlyChg < 0 Then
                ws.Cells(NewTickerCounter - 1, 11).Interior.ColorIndex = 3
             End If
             
             'Store next open price value for yearly change calculation
             OpenPrice = ws.Cells(I + 1, 3).Value
                 'ws.Cells(NewTickerCounter, 13).Value = OpenPrice  'NOT NEEDED - JUST FOR TESTING
             
            'If same ticker as previous line then add daily volume to total volume
            Else
                DailyVolume = ws.Cells(I, 7).Value
                TotalVolume = TotalVolume + DailyVolume
            End If
                         
        Next I
        
        
    'Return the stock with the Greatest % increase, greatest % decrease, and greatest total volume.

    'Create descriptors
    'ws.Cells(1, 16).Value = "Ticker"  *** is this needed? ***
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    Dim GreatIncrease As Double
    Dim GreatDecrease As Double
    Dim GreatVolume As Double
    
    'Return the stock with the Greatest % increase, greatest % decrease, and greatest total volume.
    GreatIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 17).Value = FormatPercent(GreatIncrease)
    GreatDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(3, 17).Value = FormatPercent(GreatDecrease)
    GreatVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(4, 17).Value = GreatVolume
        
        
    ' Autofit to display data
    ws.Columns("A:Q").AutoFit
        
 Next ws
 
End Sub
  

