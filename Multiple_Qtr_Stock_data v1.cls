VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData()


Dim Ticker As String
Dim YrlyChg As Double
Dim PctChg As Double
Dim TotStkVol As LongLong
Dim Rowindex As Long
Dim Columnindex As Long
Dim ws As Worksheet
Dim Start As Long
Dim Lastrow As Long
Dim find_next_ticker As Long


For Each ws In Worksheets
    Columnindex = 0
    QtrlyChg = 0
    PctChg = 0
    TotStkVol = 0
    Start = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    For Rowindex = 2 To Lastrow
    
       
        If ws.Cells(Rowindex + 1, 1).Value <> ws.Cells(Rowindex, 1).Value Then
        
            TotStkVol = TotStkVol + ws.Cells(Rowindex, 7).Value
            
            If TotStkVol = 0 Then
            
                ws.Range("I" & 2 + Columnindex).Value = Cells(Rowindex, 1).Value
                ws.Range("J" & 2 + Columnindex).Value = 0
                ws.Range("K" & 2 + Columnindex).Value = 0
                ws.Range("L" & 2 + Columnindex).Value = 0
                                    
                Else
                    
                    If ws.Cells(Start, 3) = 0 Then
                        For find_next_ticker = Start To Rowindex
                            If ws.Cells(find_next_ticker, 3).Value <> 0 Then
                                Start = find_next_ticker
                                
                                Exit For
                            
                            End If
                
                        Next find_next_ticker
                
                    End If
                                
                QtrlyChg = (ws.Cells(Rowindex, 6) - ws.Cells(Start, 3))
                PctChg = QtrlyChg / ws.Cells(Start, 3)
                
                Start = Rowindex + 1
                
                ws.Range("I" & 2 + Columnindex).Value = ws.Cells(Rowindex, 1).Value
                ws.Range("L" & 2 + Columnindex).Value = TotStkVol
                ws.Range("J" & 2 + Columnindex).Value = QtrlyChg
                ws.Range("J" & 2 + Columnindex).NumberFormat = "0.00"
                ws.Range("K" & 2 + Columnindex).Value = PctChg
                ws.Range("K" & 2 + Columnindex).NumberFormat = "0.00%"
                
                Select Case QtrlyChg
                    Case Is > 0
                        ws.Range("J" & 2 + Columnindex).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + Columnindex).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + Columnindex).Interior.ColorIndex = 0
                 End Select
                           
            End If
            
            TotStkVol = 0
            QtrlyChg = 0
            PctChg = 0
            
            Columnindex = Columnindex + 1
            
        Else
            TotStkVol = TotStkVol + ws.Cells(Rowindex, 7).Value
            
            
            
        End If
    

    Next Rowindex
    
    ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)) * 100
    ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)) * 100
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & Lastrow))

    match_max = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
    match_min = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
    match_total = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Lastrow)), ws.Range("L2:L" & Lastrow), 0)
    
    ws.Range("O2") = ws.Cells(match_max + 1, 9)
    ws.Range("O3") = ws.Cells(match_min + 1, 9)
    ws.Range("O4") = ws.Cells(match_total + 1, 9)

Next ws

    


End Sub

