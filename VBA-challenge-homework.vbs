Sub Stockmarket()


For Each ws In Worksheets


Dim ticker As String

Dim TotalVol As Double

    TotalVol = 0
    
Dim SUMTABROW As Double

    SUMTABROW = 2
    
Dim OPENROW As Double
   OPENROW = 2
   
Dim Row As Long

Dim LastRow As Long

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Max As Double
    Max = ws.Cells(2, 11).Value

Dim Min As Double
    Min = ws.Cells(2, 11).Value

Dim MaxVol As Double
    MaxVol = ws.Cells(2, 12).Value

ws.Range("I1") = "Ticker"

ws.Range("J1") = "Yearly Change"

ws.Range("K1") = "Percent Change"

ws.Range("L1") = "Total Stock Volume"

ws.Range("P1") = "Ticker"

ws.Range("Q1") = "Value"





    For Row = 2 To LastRow

      
      
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then

       

      
      
      
            ws.Cells(SUMTABROW, 9).Value = ws.Cells(Row, 1).Value
            
      
            ws.Cells(SUMTABROW, 10).Value = ws.Cells(Row, 6).Value - ws.Cells(OPENROW, 3).Value
            
            If ws.Cells(OPENROW, 3).Value = 0 Then
            
                    ws.Cells(SUMTABROW, 11) = Null
                    
            Else
            
      
                    ws.Cells(SUMTABROW, 11).Value = (ws.Cells(SUMTABROW, 10).Value / ws.Cells(OPENROW, 3).Value)
                    
            End If
        
            
      
            TotalVol = TotalVol + ws.Cells(Row, 7).Value
            
      
            ws.Cells(SUMTABROW, 12).Value = TotalVol
                 
                
                
                    
                    
                    
                    
                    
            
                    If ws.Cells(SUMTABROW, 10) < 0 Then
     
                            ws.Cells(SUMTABROW, 10).Interior.ColorIndex = 3
                            
                            
        
                    Else
    
                            ws.Cells(SUMTABROW, 10).Interior.ColorIndex = 4
        
                    End If
                    

            
      
            SUMTABROW = SUMTABROW + 1
            
           
            
            
            
            
            OPENROW = Row + 1
            
            TotalVol = 0
            
            
            
            
            
            
      Else
      
            TotalVol = TotalVol + ws.Cells(Row, 7)
      
      End If
      
      
     Next Row
     
     
     
     For Row = 2 To LastRow
     
     
     If ws.Cells(Row, 11).Value > Max Then
                    
                    Max = ws.Cells(Row, 11).Value
                    
                    
                    ws.Range("O2") = "Greatest % increase"
                    
                    ws.Range("P2") = ws.Cells(Row, 9).Value
            
                    ws.Range("Q2") = FormatPercent(Max)
                    
                    
    End If
                    
                    
                    
    If ws.Cells(Row, 11).Value < Min Then
                
                    Min = ws.Cells(Row, 11).Value
                    
                    ws.Range("O3") = "Greatest % decrease"
            
                    ws.Range("P3") = ws.Cells(Row, 9).Value
            
                    ws.Range("Q3") = FormatPercent(Min)
                    
    End If
                    
                    
   If ws.Cells(Row, 12).Value > MaxVol Then
                
                    MaxVol = ws.Cells(Row, 12).Value
                    
                    ws.Range("O4") = "Greatest Total Volume"
                    
                    ws.Range("P4") = ws.Cells(Row, 9).Value
                    
                    ws.Range("Q4") = MaxVol
                    
    End If
    
    ws.Cells(Row, 11).Value = FormatPercent(ws.Cells(Row, 11).Value)
    
    
Next Row

Next ws
     

     
    
     
End Sub
