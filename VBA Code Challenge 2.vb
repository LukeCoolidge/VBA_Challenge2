For Each Ws In Worksheets
    
Dim WorksheetName As String
Dim i As Long
Dim j As Long
Dim StockSymbol As Long
Dim EndA As Long
Dim EndI As Long
Dim PercentChange As Double
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatesttotalVolume As Double
        
        
WorksheetName = Ws.Name
          
        
'New Headers
        Ws.Range("I1").Value = "Ticker"
        Ws.Range("J1").Value = "Yearly Change"
        Ws.Range("K1").Value = "Percent Change"
        Ws.Range("L1").Value = "Total Stock Volume"
        Ws.Range("P1").Value = "Ticker"
        Ws.Range("Q1").Value = "Value"
        Ws.Range("O2").Value = "Greatest % Increase"
        Ws.Range("O3").Value = "Greatest % Decrease"
        Ws.Range("O4").Value = "Greatest Total Volume"
        
Ticker = Table

   
'Table
        Table = 2
        
'Start in Row 2
        j = 2
        
'End of row A
        EndA = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
'Loop 1
            For i = 2 To EndA
            

                If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                
'Table in "9"
                Ws.Cells(Table, 9).Value = Ws.Cells(i, 1).Value
                
'YearlyChange in "10"

                Ws.Cells(Table, 10).Value = Ws.Cells(i, 6).Value - Ws.Cells(j, 3).Value
                
'Format
                    If Ws.Cells(Table, 10).Value < 0 Then
                
'Make Red
                    Ws.Cells(Table, 10).Interior.ColorIndex = 3
                
            Else
                
'Make Green
                    Ws.Cells(Table, 10).Interior.ColorIndex = 4
                
            End If
                    
'Find The Percent Change
                    If Ws.Cells(j, 3).Value <> 0 Then
                    
                    PercentChange = ((Ws.Cells(i, 6).Value - Ws.Cells(j, 3).Value) / Ws.Cells(j, 3).Value)
                    

                    Ws.Cells(Table, 11).Value = Format(PercentChange, "Percent")
                    
            Else
                    
                    Ws.Cells(Table, 11).Value = Format(0, "Percent")
                    
            End If
                    
            
                Ws.Cells(Table, 12).Value = WorksheetFunction.Sum(Range(Ws.Cells(j, 7), Ws.Cells(i, 7)))
                
'Add 1 to Table
                Table = Table + 1
                
'Next begining Row
                j = i + 1
                
            End If
            
            Next i
            

            
    Next Ws