Attribute VB_Name = "Module1"
Sub StockSummaryTable()

For Each ws In Worksheets

       
        'Variables
        
        Dim Stock As String
        Dim ticker As String
        Dim Day As Date
        Dim DayOpen As Double
        Dim DayHigh As Double
        Dim DayLow As Double
        Dim DayClose As Double
        Dim DayVol As Double
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVol As Double
        Dim Vol As Double
        Dim DataLastRow As Long
        Dim SummaryLastRow As Long
        Dim SummaryLastRow2 As Long
        Dim SummaryTicker As String
        Dim SummaryRowCount As Integer
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVol As Double
        Dim RowI As Integer
        Dim RowD As Integer
        Dim RowV As Integer
        Dim TickerV As String
        Dim TickerI As String
        Dim TickerD As String
        
        
        
         
                
        'How to find the last row for my loop (Source https://excelmacromastery.com/vba-for-loop/#VBA_For_Loop_Example_3)
        
        DataLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
        
        ' Create Headers For Summary Table
          
          ws.Range("I1").Value = "Ticker"
          ws.Range("J1").Value = "Yearly Change"
          ws.Range("K1").Value = "Percent Change"
          ws.Range("L1").Value = "Total Stock Volume"
          ws.Range("N2").Value = "Greatest % Increase"
          ws.Range("N3").Value = "Greatest % Decrease"
          ws.Range("N4").Value = "Greatest Total Volumne"
          ws.Range("O1").Value = "Ticker"
          ws.Range("P1").Value = "Value"
          
                     
       
       'Set Variable Values
       
       TotalVol = 0
       DayOpen = ws.Cells(2, 3).Value
       SummaryRowCount = 2
       GreatestVol = 0
       GreatestIncrease = 0
       RowCount = 2
       
       'Create loop to move through stock data
        
        For i = 2 To DataLastRow
        
        
            'Add vol to total
            
            TotalVol = TotalVol + ws.Cells(i, 7).Value
            
        
            'finding the last entry for that stock in the list
            
                If ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value Then
                
                    'Set Ticker name
                    ticker = ws.Cells(i, 1).Value
                    
                    DayClose = ws.Cells(i, 6).Value
                    
                    'compare the change in first open value to last close value
                    YearlyChange = DayClose - DayOpen
                    
                    'Calculate the % change in price for the year
                    If DayOpen <> 0 Then
                        PercentChange = (YearlyChange / DayOpen) * 100
                    Else
                        DayOpen = ActiveCell.Offset(1, 0).Select
                        PercentChange = (YearlyChange / DayOpen) * 100
                        
                    End If
                    
                    'add values to summary table
                    
                    ws.Cells(SummaryRowCount, 9).Value = ticker
                    ws.Cells(SummaryRowCount, 10).Value = YearlyChange
                    
                     'Format Cells so + change is green and - change is read
                    
                        If ws.Cells(SummaryRowCount, 10).Value > 0 Then
                            
                            ws.Cells(SummaryRowCount, 10).Interior.ColorIndex = 4
                            
                        ElseIf ws.Cells(SummaryRowCount, 10).Value < 0 Then
                        
                            
                            ws.Cells(SummaryRowCount, 10).Interior.ColorIndex = 3
                            
                        End If
                            
                    
                    ws.Cells(SummaryRowCount, 11).Value = PercentChange
                    ws.Cells(SummaryRowCount, 11).NumberFormat = "0.00%"
                    ws.Cells(SummaryRowCount, 12).Value = TotalVol
                    
                    
                    'Reset DayOpen to Next stock
                    
                    DayOpen = ws.Cells((i + 1), 3).Value
                    
                
                    'update summary row count to add 1 row to the summary table
                    
                    SummaryRowCount = SummaryRowCount + 1
                    
                    
                    TotalVol = 0
                
                End If
                    
        Next i
            
                  
        'Find length of Summary Tabel
        
        SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
               
        
    'Create second summary table
        
   'Define Variables for second summary table
        
    GreatestDecrease = ws.Cells(2, 11).Value
    GreatestIncrease = ws.Cells(2, 11).Value
    GreatestVol = ws.Cells(2, 12).Value
    TickerD = ws.Cells(2, 9).Value
    TickerI = ws.Cells(2, 9).Value
    TickerV = ws.Cells(2, 9).Value
    RowV = 3
    RowI = 3
    RowD = 3
    
    
   'create a loop to move through the summary stock value for the second summary table
   
   
   For i = 2 To SummaryLastRow
   
        
        'Find the biggest volume
        If ws.Cells(RowV, 12).Value <= ws.Cells(i, 12).Value Then
        
            
            TickerV = ws.Cells((i), 9).Value
            
            GreatestVol = ws.Cells((i), 12).Value
            
            RowV = i
            
        End If
        
        
        'Find biggest increase
        If ws.Cells(RowI, 11).Value <= ws.Cells(i, 11).Value Then
        
            TickerI = ws.Cells(i, 9).Value
            
            GreatestIncrease = ws.Cells(i, 11).Value
            
            RowI = i
            
        End If
        
        
        'find biggest decrease
        If ws.Cells(RowD, 11).Value >= ws.Cells(i, 11).Value Then
        
            TickerD = ws.Cells(i, 9).Value
            GreatestDecrease = ws.Cells(i, 11).Value
            
            RowD = i
            
        End If
        
        
        
    Next i
    
    
    
    'Add values to the second table
    
    ws.Cells(4, 15).Value = TickerV
    ws.Cells(4, 16).Value = GreatestVol
    ws.Cells(2, 15).Value = TickerI
    ws.Cells(2, 16).Value = GreatestIncrease
    ws.Cells(3, 15).Value = TickerD
    ws.Cells(3, 16).Value = GreatestDecrease
   
         
  'Format Columns
        
  ws.Cells.EntireColumn.AutoFit
  

Next ws


End Sub

