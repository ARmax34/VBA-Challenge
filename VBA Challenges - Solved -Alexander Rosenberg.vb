Sub StockInfo1():

'looping throught the different worksheets
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    'creating table header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'creating second table labels
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
        'creating variables
        
            'ticker name columns
            Dim stockTickerColoum As Integer
                stockTickerColoum = 1
            'Ticker name writer coloum
            Dim tickerNameWriterColoum As Integer
                tickerNameWriterColoum = 9
            'Ticker Name Writter counter
            Dim tickerNameWritterCounter As Integer
                tickerNameWritterCounter = 2
                    
    
            'volume storage variable
            Dim vol As Double
                vol = ws.Cells(2, 7)
    
            'variable to record at the opening of each peroiod
            Dim changeOpen As Single
                changeOpen = ws.Cells(2, 3)
                    
                    
            'variable to record at the opening of each peroiod
            Dim changeClose As Single
                changeClose = 0
    
        'create row counter
        Dim n As Double
        'For n = 1 To Rows.Count
        'Next n
        'MsgBox (n)
        
        n = Range("A2", Range("A1").End(xlDown)).Rows.Count
        'MsgBox (n)
        
        
        ' Loop through rows in the column
        For i = 2 To n
        
            'if statement for when the ticker matches
            If ws.Cells(i + 1, stockTickerColoum).Value = ws.Cells(i, stockTickerColoum).Value Then
                
      
                'Adding the amounts to the variables
                vol = vol + ws.Cells(i + 1, 7).Value
                
                
            'elseif for when the ticker does not matche
            ElseIf ws.Cells(i + 1, stockTickerColoum).Value <> ws.Cells(i, stockTickerColoum).Value Then
            
                'Testing box
                'MsgBox (Cells(i, stockTickerColoum).Value & " and then " & Cells(i + 1, stockTickerColoum).Value)
                
                    'Entering in the information into the table
                    ws.Cells(tickerNameWritterCounter, tickerNameWriterColoum).Value = ws.Cells(i, stockTickerColoum)
    
                
                'capturing the closing price data
                changeClose = ws.Cells(i, 5 + stockTickerColoum).Value
                
                
                'calculating the yearly change
                ws.Cells(tickerNameWritterCounter, 1 + tickerNameWriterColoum) = changeClose - changeOpen
                
                'Calculating the percentage change
                ws.Cells(tickerNameWritterCounter, 2 + tickerNameWriterColoum) = (changeClose / changeOpen) - 1
                   
                    
                
                'Formating cells
                
                    'Changing the Yearly change coloum To be formated for display
                    ws.Cells(tickerNameWritterCounter, 1 + tickerNameWriterColoum).NumberFormat = "0.00;-0.00"
                    

                            
             

                    'Changing the Peccentage coloum To be formated for display
                    ws.Cells(tickerNameWritterCounter, 2 + tickerNameWriterColoum).NumberFormat = "0.00%;-0.00%"
                
                
                
                
                'new valus of change open
                changeOpen = ws.Cells(i + 1, 3)
        
        
                    'volume cells data entering
                    ws.Cells(tickerNameWritterCounter, 3 + tickerNameWriterColoum).Value = vol
        
                    vol = ws.Cells(i + 1, 7)
                
                
                'moving to next row on table
                tickerNameWritterCounter = tickerNameWritterCounter + 1
                    
                 
           End If

        
        Next i
        
        
        
        'filling secondary table
        Dim tableRange As Integer
        tableRange = Range("I1", Range("I1").End(xlDown)).Rows.Count
        'MsgBox (tableRange)
        
        'Greatest Variables
        Dim greatestIncrease As Double
        greatestIncrease = -10
            
        Dim greatestDecrease As Double
        greatestDecrease = 10
        
        Dim greatestVolume As Double
        greatestVolume = -1
        
        
            For u = 2 To tableRange

            
                            
                'GreatestIncrease IF
                If ws.Cells(u, 11).Value >= greatestIncrease Then
                    
                    greatestIncrease = ws.Cells(u, 11).Value
                    
                        'for the ticker
                        ws.Cells(2, 17).Value = ws.Cells(u, 11).Value
                    
                        'for the value
                        ws.Cells(2, 16).Value = ws.Cells(u, 9).Value

                End If
                
                'GreatestDecrease IF
                If ws.Cells(u, 11).Value <= greatestDecrease Then
                        greatestDecrease = ws.Cells(u, 11).Value
                    
                        'for the ticker
                        ws.Cells(3, 17).Value = ws.Cells(u, 11).Value
                    
                        'for the value
                        ws.Cells(3, 16).Value = ws.Cells(u, 9).Value

                End If
                   
                
                'GreatestVolume IF
                If ws.Cells(u, 12).Value >= greatestVolume Then
                        greatestVolume = ws.Cells(u, 12).Value
                    
                        'for the ticker
                        ws.Cells(4, 17).Value = ws.Cells(u, 12).Value
                    
                        'for the value
                        ws.Cells(4, 16).Value = ws.Cells(u, 9).Value
                        
                End If
                
                
                
                
            'Formating cells
                
                'secondary table
                    'Changing the Peccentage coloum To be formated for display
                        ws.Cells(2, 17).NumberFormat = "0.00%;[red]-0.00%"
           
                    'Changing the Peccentage coloum To be formated for display
                        ws.Cells(3, 17).NumberFormat = "0.00%;[red]-0.00%"
                
                
                'primary table
                    'adding red to the cells
                        If ws.Cells(u, 10) < 0 Then
                            ws.Cells(u, 10).Interior.ColorIndex = 3
                            
                    'adding green to the cells
                        ElseIf ws.Cells(u, 10) > 0 Then
                            ws.Cells(u, 10).Interior.ColorIndex = 4
                
                        End If
                
            Next u

    Next ws


End Sub



